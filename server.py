import os
import time
import argparse
import asyncio
import concurrent.futures
from typing import Dict, Any
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse, FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import uvicorn
import cv2
import numpy as np

# Thread pool for running blocking YOLO inference without freezing the event loop
_thread_pool = concurrent.futures.ThreadPoolExecutor(max_workers=2)

# Import the monitor logic
import workplace_safety_monitor
from workplace_safety_monitor import get_monitor

app = FastAPI(title="Workplace Safety Monitor API")

# CORS middleware for development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global state — initialized at module level to avoid NameError on /health
_current_source: str = "test_videos/veo3_construction.mp4"
_monitor_lock = asyncio.Lock()

latest_stats: Dict[str, Any] = {
    "frame_id": 0,
    "timestamp": 0,
    "total_persons": 0,
    "compliant": 0,
    "violations": 0,
    "workers": []
}

# Sentinel used to safely detect StopIteration across thread boundaries
_SENTINEL = object()


def _next_frame(gen):
    """
    Safe wrapper around next() for use with run_in_executor.
    StopIteration cannot propagate through asyncio coroutines (PEP 479),
    so we convert it to a sentinel value instead.
    """
    try:
        return next(gen)
    except StopIteration:
        return _SENTINEL


def get_default_args(source: str = "0") -> argparse.Namespace:
    """Helper to create namespace with requested defaults."""
    return argparse.Namespace(
        source=source,
        person_weights='yolo12n.pt',
        person_conf=0.45,
        ppe_weights='best.pt',
        ppe_onnx=None,
        class_names=['helmet', 'vest'],
        input_size=640,
        conf_helmet=0.5,
        conf_vest=0.5,
        nms_iou=0.45,
        min_box_area=900,
        max_aspect_ratio=3.5,
        head_iou_gate=0.10,
        torso_iou_gate=0.15,
        ppe_inside_person_frac=0.35,
        temporal_window=7,
        track_iou=0.3,
        track_max_age=20,
        track_edge_cleanup=False,
        save_vis='',
        eval_root=''
    )


def create_no_signal_frame() -> bytes:
    """Create a black frame with 'NO SIGNAL' text for stream placeholder."""
    frame = np.zeros((480, 640, 3), dtype=np.uint8)
    cv2.putText(frame, "NO SIGNAL", (180, 240),
                cv2.FONT_HERSHEY_SIMPLEX, 1.5, (0, 0, 255), 3)
    _, buffer = cv2.imencode('.jpg', frame)
    return buffer.tobytes()


def init_monitor(source: str):
    """Initialize or reinitialize the monitor singleton (not thread-safe alone — use _monitor_lock)."""
    global _current_source
    _current_source = source
    if workplace_safety_monitor._monitor_instance:
        workplace_safety_monitor._monitor_instance.release()
    workplace_safety_monitor._monitor_instance = None
    return get_monitor(get_default_args(source))


@app.on_event("startup")
async def startup_event():
    """Initialize monitor with default video on startup."""
    default = "test_videos/veo3_construction.mp4"
    if not os.path.exists(default):
        print(f"[WARN] Default source not found: {default}. Dashboard will show NO SIGNAL until a source is set.")
        return
    init_monitor(default)


@app.on_event("shutdown")
async def shutdown_event():
    """Gracefully release resources on shutdown."""
    _thread_pool.shutdown(wait=False)
    if workplace_safety_monitor._monitor_instance:
        workplace_safety_monitor._monitor_instance.release()
    print("[INFO] Server shut down cleanly.")


# Create static directory if missing
static_dir = os.path.join(os.path.dirname(__file__), "static")
if not os.path.exists(static_dir):
    os.makedirs(static_dir)


@app.get("/")
async def get_index():
    """Serve the dashboard frontend."""
    index_path = os.path.join(static_dir, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return JSONResponse({"error": "index.html not found in static/"}, status_code=404)


async def frame_generator():
    """
    MJPEG stream generator.
    Runs blocking YOLO inference in a thread pool so the asyncio event loop
    stays free to handle concurrent requests (e.g. source switching).
    Uses _SENTINEL to safely detect generator exhaustion across thread boundaries.
    """
    global latest_stats
    no_signal_frame = create_no_signal_frame()
    loop = asyncio.get_running_loop()  # get_event_loop() is deprecated in 3.10+

    while True:
        mon = workplace_safety_monitor._monitor_instance
        if mon is None:
            yield (b'--frame\r\n'
                   b'Content-Type: image/jpeg\r\n\r\n' + no_signal_frame + b'\r\n')
            await asyncio.sleep(0.5)
            continue

        try:
            gen = mon.get_frames()
            while True:
                # Run blocking frame-read + inference in the thread pool
                result = await loop.run_in_executor(_thread_pool, _next_frame, gen)

                # Video ended (StopIteration was converted to sentinel by _next_frame)
                if result is _SENTINEL:
                    break

                jpeg_bytes, stats = result
                latest_stats = stats

                # Source was changed — pick up the new monitor instance
                if workplace_safety_monitor._monitor_instance is not mon:
                    break

                yield (b'--frame\r\n'
                       b'Content-Type: image/jpeg\r\n\r\n' + jpeg_bytes + b'\r\n')

            # Brief pause before restarting (video loop or source switch)
            await asyncio.sleep(0.1)

        except Exception as e:
            print(f"[ERROR] Streaming error: {e}")
            yield (b'--frame\r\n'
                   b'Content-Type: image/jpeg\r\n\r\n' + no_signal_frame + b'\r\n')
            await asyncio.sleep(1)


@app.get("/stream")
async def stream_video():
    """Endpoint for MJPEG video stream."""
    return StreamingResponse(
        frame_generator(),
        media_type="multipart/x-mixed-replace; boundary=frame"
    )


@app.get("/stats")
async def get_current_stats():
    """Return the most recent detection statistics."""
    return latest_stats


@app.post("/source")
async def change_source(payload: Dict[str, str]):
    """Update the monitoring source and reinitialize the monitor."""
    source = payload.get("source")
    if source is None:
        raise HTTPException(status_code=400, detail="Missing 'source' in request body")

    # Validate source type
    is_valid = False
    if source.isdigit():
        is_valid = True
    elif source.lower().startswith("rtsp://") and len(source) > 10:
        # Require something after rtsp:// to avoid bare invalid URLs
        is_valid = True
    elif os.path.exists(source):
        is_valid = True

    if not is_valid:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid source '{source}'. Must be a camera index (e.g. 0), "
                   "a valid RTSP URL (rtsp://...), or an existing local file path."
        )

    # Lock ensures concurrent /source requests don't corrupt monitor state
    async with _monitor_lock:
        init_monitor(source)

    return {"status": "success", "message": f"Source updated to {source}"}


@app.get("/health")
async def health_check():
    """Health endpoint showing system status."""
    mon = workplace_safety_monitor._monitor_instance
    is_streaming = bool(mon and mon.cap and mon.cap.isOpened())

    return {
        "status": "ok",
        "streaming": is_streaming,
        "source": _current_source,
        "timestamp": time.time()
    }


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
