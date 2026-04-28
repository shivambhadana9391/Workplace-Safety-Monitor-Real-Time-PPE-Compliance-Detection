# 🏗️ Workplace Safety Monitor - Comprehensive Project Audit & Report

This document provides a complete technical audit and formal report for the Workplace Safety Monitor AI system.

---

## 📂 Part 1: Codebase Audit & File Tree

### Project Structure

```text
ppe-safety-detection-ai/
├── dataset/                        # Labeled data for model training/validation
│   └── safety-Helmet-Reflective-Jacket/
│       ├── test/                   # Test images and YOLO-format labels
│       ├── train/                  # Training images and YOLO-format labels
│       ├── valid/                  # Validation images and YOLO-format labels
│       └── data.yaml               # Dataset configuration (classes, paths)
├── output/                         # Results from processed video/streams
│   ├── frame_*.jpg                 # Individual annotated frames
│   └── output.mp4                  # Final annotated video output
├── test_videos/                    # Sample footage for testing
│   └── veo3_construction.mp4       # Sample construction site video
├── best.pt                         # Custom YOLO weights for Helmet & Vest detection
├── yolo12n.pt                      # Pre-trained weights for Person detection
├── workplace_safety_monitor.py     # Main application logic
├── requirements.txt                # Python dependencies
├── README.md                       # System documentation and usage guide
└── output.gif                      # Visual demo of the system
```

---

### File Descriptions

*   **`workplace_safety_monitor.py`**: The central script of the project. It handles video ingestion (files, webcams, or IP cameras), runs two-stage detection (Persons → PPE), associates equipment to specific individuals using anatomical region logic, and applies temporal smoothing to ensure stable, reliable alerts.
*   **`requirements.txt`**: Lists the necessary Python packages, primarily `ultralytics` (for YOLOv8/v11/v12 support), `opencv-python` (for video I/O and visualization), and `numpy`.
*   **`best.pt`**: The specialized YOLO model weights. This model was trained to identify safety helmets and reflective vests with high precision.
*   **`yolo12n.pt`**: A lightweight, high-speed model used specifically to detect people. By detecting people first, the system can more accurately attribute PPE to the correct worker.
*   **`dataset/`**: A structured repository of images and text labels following the YOLOv8 format. It is used to evaluate the model's accuracy and could be used for further fine-tuning.
*   **`test_videos/`**: Contains raw video samples (like `veo3_construction.mp4`) that serve as benchmarks for the system's performance in varied lighting and environments.
*   **`output/`**: The directory where the system exports its findings. It captures both individual frame "snapshots" and a compiled MP4 video with real-time bounding boxes and compliance status.
*   **`README.md`**: A detailed guide containing installation steps, CLI arguments, technical architecture details, and troubleshooting tips.
*   **`output.gif`**: A high-level visual summary showing the system successfully tracking multiple workers and flagging compliance status (Green for safe, Red for non-compliant).

---

## 📑 Part 2: Technical Project Report

### 1. Executive Summary
The **Workplace Safety Monitor** is a high-performance computer vision solution designed to automate the monitoring of Personal Protective Equipment (PPE) in industrial and construction environments. By leveraging state-of-the-art YOLO (You Only Look Once) models, the system identifies personnel and verifies their compliance with safety protocols (helmets and safety vests) in real-time.

---

### 2. Objectives
*   **Automated Compliance**: Replace manual safety checks with 24/7 AI monitoring.
*   **Accurate Association**: Ensure PPE is correctly attributed to the individual wearing it.
*   **Real-time Response**: Provide instantaneous visual feedback and logging for safety violations.
*   **Versatile Integration**: Support various hardware inputs from standard webcams to industrial IP cameras.

---

### 3. Technical Stack
*   **Core Logic**: Python 3.8+
*   **Computer Vision**: OpenCV (Open Source Computer Vision Library)
*   **Deep Learning**: Ultralytics YOLOv8/v12 Framework
*   **Inference Backends**: PyTorch (.pt) and ONNX for cross-platform compatibility
*   **Data Handling**: NumPy for optimized tensor operations

---

### 4. System Architecture

#### A. Two-Stage Detection Pipeline
The system avoids the common pitfall of single-pass detection by using a specialized two-stage process:
1.  **Person Detection**: Identifies humans in the frame using high-speed models (`yolo12n.pt`).
2.  **PPE Detection**: A secondary, custom-trained model (`best.pt`) focuses exclusively on safety gear within the identified human boundaries.

#### B. Anatomical Region Logic
To prevent false compliance (e.g., a helmet hanging on a wall behind a worker), the system implements **Anatomical Mapping**:
*   **Head Gate**: Helmets are only validated if they overlap with the top 25-30% of the detected person's bounding box.
*   **Torso Gate**: Safety vests must reside in the central 50% of the person's bounding box.

#### C. Temporal Smoothing & Tracking
The system uses an **IoU (Intersection over Union) Tracker** to maintain the identity of workers across frames. A **Temporal Smoother** then analyzes a rolling window of detections (default: 7 frames) to confirm compliance, effectively filtering out momentary occlusions or model glitches.

---

### 5. Implementation Methodology

| Component | Description |
| :--- | :--- |
| **Detection Backend** | Supports both `.pt` (PyTorch) for high accuracy and `.onnx` for high-speed edge deployment. |
| **Association Algorithm** | Greedy one-to-one matching based on spatial overlap and containment scores. |
| **Edge Cleanup** | Automatically identifies and removes tracking boxes when personnel exit the camera frame. |
| **Visualization** | Dynamic bounding boxes: **Green** (Compliant), **Red** (Violation), and **Blue** (PPE detection). |

---

### 6. Usage & Deployment

The system is controlled via a robust CLI (Command Line Interface):

```bash
# Standard Real-time Monitoring
python workplace_safety_monitor.py --source 0 --ppe-weights best.pt

# Industrial RTSP Stream Integration
python workplace_safety_monitor.py --source "rtsp://camera_ip" --save-vis output/
```

---

### 7. Performance Evaluation
The system includes a dedicated `--eval-root` mode to benchmark accuracy against labeled datasets. It reports:
*   **Precision (P)**: Accuracy of compliance alerts.
*   **Recall (R)**: Ability to find all personnel/PPE in the frame.
*   **Threshold Suggestions**: Automated recommendations to tune the system for specific lighting or distances.

---

### 8. Conclusion & Future Scope
The Workplace Safety Monitor provides a robust foundation for industrial safety. Future enhancements include:
*   **Multi-Camera Support**: Unified dashboard for entire facility monitoring.
*   **Additional PPE**: Support for safety goggles, gloves, and specialized footwear.
*   **Database Integration**: Automated logging of safety violations into enterprise SQL/Cloud databases.

---
**Disclaimer**: This AI system is an assistive tool and should be used alongside existing safety protocols and human oversight.
