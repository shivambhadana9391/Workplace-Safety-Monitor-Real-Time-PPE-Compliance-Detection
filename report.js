const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  ExternalHyperlink, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

const BORDER = { style: BorderStyle.SINGLE, size: 8, color: "000000" };
const ALL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDERS = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function cell(text, bold = false, width = 2340) {
  return new TableCell({
    borders: ALL_BORDERS,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, bold, font: "Arial", size: 22 })]
    })]
  });
}

function infoTable(rows) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2340, 2340, 2340, 2340],
    rows: rows.map(r => new TableRow({
      children: r.map((c, i) => cell(c.text, c.bold, 2340))
    }))
  });
}

function twoColTable(rows, col1w = 3120, col2w = 6240) {
  const border = BORDER;
  const borders = ALL_BORDERS;
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [col1w, col2w],
    rows: rows.map(r => new TableRow({
      children: [
        new TableCell({
          borders, width: { size: col1w, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: r[0], bold: r[0].startsWith('**') || true, font: "Arial", size: 22 })] })]
        }),
        new TableCell({
          borders, width: { size: col2w, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: r[1], font: "Arial", size: 22 })] })]
        })
      ]
    }))
  });
}

function twoColTableBold(rows, col1w = 3120, col2w = 6240) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [col1w, col2w],
    rows: rows.map((r, ri) => new TableRow({
      children: [
        new TableCell({
          borders: ALL_BORDERS,
          width: { size: col1w, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          shading: ri === 0 ? { fill: "EEEEEE", type: ShadingType.CLEAR } : undefined,
          children: [new Paragraph({ children: [new TextRun({ text: r[0], bold: ri === 0, font: "Arial", size: 22 })] })]
        }),
        new TableCell({
          borders: ALL_BORDERS,
          width: { size: col2w, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          shading: ri === 0 ? { fill: "EEEEEE", type: ShadingType.CLEAR } : undefined,
          children: [new Paragraph({ children: [new TextRun({ text: r[1], bold: ri === 0, font: "Arial", size: 22 })] })]
        })
      ]
    }))
  });
}

function threeColTable(headers, rows) {
  const w = [3120, 3120, 3120];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: w,
    rows: [
      new TableRow({
        children: headers.map((h, i) => new TableCell({
          borders: ALL_BORDERS,
          width: { size: w[i], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          shading: { fill: "EEEEEE", type: ShadingType.CLEAR },
          children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, font: "Arial", size: 22 })] })]
        }))
      }),
      ...rows.map(r => new TableRow({
        children: r.map((c, i) => new TableCell({
          borders: ALL_BORDERS,
          width: { size: w[i], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: c, font: "Arial", size: 22 })] })]
        }))
      }))
    ]
  });
}

function fourColTable(headers, rows) {
  const w = [2340, 2340, 2340, 2340];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: w,
    rows: [
      new TableRow({
        children: headers.map((h, i) => new TableCell({
          borders: ALL_BORDERS,
          width: { size: w[i], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          shading: { fill: "EEEEEE", type: ShadingType.CLEAR },
          children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, font: "Arial", size: 22 })] })]
        }))
      }),
      ...rows.map(r => new TableRow({
        children: r.map((c, i) => new TableCell({
          borders: ALL_BORDERS,
          width: { size: w[i], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: c, font: "Arial", size: 22 })] })]
        }))
      }))
    ]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, font: "Arial", size: 46 })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", size: 32 })]
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", size: 26 })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 276, lineRule: "auto" },
    children: [new TextRun({ text, font: "Arial", size: 22, ...opts })]
  });
}

function pBold(label, rest) {
  return new Paragraph({
    spacing: { after: 120, line: 276, lineRule: "auto" },
    children: [
      new TextRun({ text: label, bold: true, font: "Arial", size: 22 }),
      new TextRun({ text: rest, font: "Arial", size: 22 })
    ]
  });
}

function bullet(text, bold_prefix = null) {
  const children = bold_prefix
    ? [new TextRun({ text: bold_prefix, bold: true, font: "Arial", size: 22 }),
       new TextRun({ text, font: "Arial", size: 22 })]
    : [new TextRun({ text, font: "Arial", size: 22 })];
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children
  });
}

function codeBlock(lines) {
  return lines.map(line => new Paragraph({
    spacing: { after: 0, line: 240, lineRule: "auto" },
    indent: { left: 720 },
    children: [new TextRun({ text: line, font: "Courier New", size: 18 })]
  }));
}

function sp(after = 120) {
  return new Paragraph({ spacing: { after }, children: [new TextRun("")] });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function link(text, url) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new ExternalHyperlink({
      children: [new TextRun({ text, style: "Hyperlink", font: "Arial", size: 22 })],
      link: url
    })]
  });
}

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 46, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 400, after: 120, line: 276, lineRule: "auto" }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, font: "Arial" },
        paragraph: { spacing: { before: 360, after: 120, line: 276, lineRule: "auto" }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, font: "Arial" },
        paragraph: { spacing: { before: 320, after: 80, line: 276, lineRule: "auto" }, outlineLevel: 2 }
      },
    ]
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "AAAAAA", space: 1 } },
          children: [new TextRun({ text: "AI-Powered PPE Safety Detection System  |  INT428 Project Report", font: "Arial", size: 18, color: "666666" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: "AAAAAA", space: 1 } },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            new TextRun({ text: "Lovely Professional University  |  B.Tech CSE", font: "Arial", size: 18, color: "666666" }),
            new TextRun({ text: "\t", font: "Arial", size: 18 }),
            new TextRun({ text: "Page ", font: "Arial", size: 18, color: "666666" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "666666" }),
          ]
        })]
      })
    },
    children: [
      // ─── COVER PAGE ───────────────────────────────────────────
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: "AI-Powered PPE Safety Detection System", bold: true, font: "Arial", size: 46 })]
      }),
      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: "A Comprehensive Project Report", bold: true, font: "Arial", size: 26, color: "000000" })]
      }),
      sp(240),
      new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Submitted by:", bold: true, font: "Arial", size: 22 })] }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 2340, 2340, 2340],
        rows: [
          new TableRow({
            children: [
              { text: "Name :", bold: true }, { text: "Roll No. :", bold: true },
              { text: "Registration No. :", bold: true }, { text: "Branch / Semester :", bold: true }
            ].map(c => new TableCell({
              borders: ALL_BORDERS,
              width: { size: 2340, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: c.text, bold: c.bold, font: "Arial", size: 22 })] })]
            }))
          }),
          new TableRow({
            children: [
              { text: "[Your Full Name]" }, { text: "[Roll No.]" },
              { text: "[Registration No.]" }, { text: "B.Tech CSE / 4th Sem" }
            ].map(c => new TableCell({
              borders: ALL_BORDERS,
              width: { size: 2340, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: c.text, font: "Arial", size: 22 })] })]
            }))
          })
        ]
      }),
      sp(120),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({ text: "GitHub Repository: ", bold: true, font: "Arial", size: 22 }),
          new ExternalHyperlink({
            children: [new TextRun({ text: "https://github.com/prodbykosta/ppe-safety-detection-ai", style: "Hyperlink", font: "Arial", size: 22 })],
            link: "https://github.com/prodbykosta/ppe-safety-detection-ai"
          })
        ]
      }),
      sp(120),
      new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Submitted to:", bold: true, font: "Arial", size: 22 })] }),
      bullet("Faculty Guide: Dr. Jimmy Singla"),
      bullet("Institution: Lovely Professional University"),
      sp(120),
      pageBreak(),

      // ─── DECLARATION ──────────────────────────────────────────
      h2("Declaration"),
      p('I hereby declare that the project entitled "AI-Powered PPE Safety Detection System" submitted in partial fulfillment of the requirements for the INT428 assessment is an authentic record of my own work carried out under the guidance of Dr. Jimmy Singla. The matter embodied in this report has not been submitted by me for the award of any other degree or diploma to this or any other university/institute.'),
      sp(240),
      p("Date:  ___________________________"),
      sp(120),
      p("Signature:  ___________________________"),
      sp(120),
      pageBreak(),

      // ─── ACKNOWLEDGMENTS ──────────────────────────────────────
      h2("Acknowledgments"),
      p("I would like to express my sincere gratitude to my project guide, Dr. Jimmy Singla, for their invaluable guidance, encouragement, and consistent support throughout the development of this project. I also extend my appreciation to the Department of Computer Science and Engineering at Lovely Professional University for providing the infrastructure and academic environment that made this work possible."),
      sp(120),
      p("Special thanks to the open-source community behind Ultralytics YOLO, OpenCV, and the Kaggle dataset contributors whose publicly available resources formed the foundation of this system."),
      pageBreak(),

      // ─── ABSTRACT ─────────────────────────────────────────────
      h2("Abstract"),
      p("Workplace safety incidents caused by the absence or improper use of Personal Protective Equipment (PPE) remain a critical challenge across construction, manufacturing, and industrial sectors globally. This project presents an AI-Powered PPE Safety Detection System designed to provide autonomous, real-time monitoring of worker compliance with mandatory safety equipment requirements."),
      sp(120),
      p("The system employs a dual-model computer vision architecture. A primary YOLO-based person detection model (yolo12n.pt) identifies individual workers within a video frame, while a custom-trained PPE detection model (best.pt), fine-tuned on the Construction Site Safety Image Dataset by Snehil Sanyal, performs granular helmet and safety vest classification. A proprietary intelligent association engine uses anatomical region analysis to correctly assign detected PPE items to their corresponding workers, eliminating false associations in crowded multi-person scenes."),
      sp(120),
      p("The system supports multiple input sources including live webcam feeds, IP camera streams, and pre-recorded video files, and outputs annotated MP4 recordings that can be used for audit trails, worker training, and compliance reporting. Temporal smoothing across consecutive frames significantly reduces false positive detections, resulting in reliable, production-grade outputs. This project bridges the gap between advanced computer vision research and practical, deployable workplace safety technology."),
      pageBreak(),

      // ─── TABLE OF CONTENTS ────────────────────────────────────
      h2("Table of Contents"),
      bullet("Chapter 1: Introduction"),
      bullet("Chapter 2: Literature Review"),
      bullet("Chapter 3: Theoretical Framework"),
      bullet("Chapter 4: System Architecture & Data Pipeline"),
      bullet("Chapter 5: Implementation & Technical Details"),
      bullet("Chapter 6: Results and Discussion"),
      bullet("Chapter 7: Societal Impact and Commercial Value"),
      bullet("Chapter 8: Conclusion and Future Scope"),
      bullet("References"),
      bullet("Appendix: LLM Usage Report (Annexure 3b Sections)"),
      pageBreak(),

      // ─── CHAPTER 1 ────────────────────────────────────────────
      h1("Chapter 1: Introduction"),

      h3("1.1 Background of the Problem"),
      p("Construction sites, manufacturing floors, and industrial facilities rank among the most hazardous working environments in the world. According to occupational safety literature, the failure to wear mandatory Personal Protective Equipment (PPE) such as safety helmets and high-visibility vests is a leading contributing factor to preventable workplace injuries and fatalities. Traditional safety enforcement relies entirely on human supervisors conducting periodic, manual walkthroughs, a process that is inherently intermittent, subjective, and unscalable."),
      sp(120),
      p("As workforce sizes grow and facility footprints expand, a single safety officer cannot maintain continuous 360-degree visibility across an entire site. Violations can occur and persist for significant durations without detection, creating windows of unacceptable risk. The challenge, therefore, is to develop a system that can perform continuous, objective, and automated PPE compliance monitoring without human bottlenecks."),

      h3("1.2 Motivation"),
      p("The proliferation of affordable, high-resolution camera infrastructure in modern industrial facilities, combined with dramatic advances in real-time deep learning inference, creates a unique technological opportunity. YOLO (You Only Look Once) architectures have achieved near-human-level accuracy in object detection tasks while running at speeds sufficient for live video analysis. This project is motivated by the belief that these technologies can be combined into a deployable system that acts as an always-on, tireless safety compliance monitor, augmenting human supervisors rather than replacing them."),
      sp(120),
      p("A key design motivation is the practical usability of outputs. The system is engineered not just to detect violations but to produce annotated video recordings that can serve as legally defensible audit evidence and as educational material for worker safety training programs."),

      h3("1.3 Project Objectives"),
      p("The primary objectives of this project are:"),
      bullet("Real-Time Detection: To deploy a multi-class object detection pipeline capable of identifying persons, helmets, and safety vests in live video at operationally acceptable frame rates."),
      bullet("Intelligent PPE-to-Person Association: To implement an anatomical region analysis engine that correctly attributes PPE items to specific workers, resolving the multi-person assignment problem inherent in crowded scenes."),
      bullet("Multi-Source Input Support: To engineer the system to accept diverse inputs including USB webcams (index-based), IP camera RTSP streams, and pre-recorded video files."),
      bullet("Annotated Video Output: To generate MP4 video recordings with overlaid detection bounding boxes, worker IDs, and compliance status labels for audit and training use."),
      bullet("Robustness via Temporal Smoothing: To implement frame-to-frame consistency checks that suppress transient false positive detections and improve overall system reliability."),

      h3("1.4 Scope and Limitations"),
      p("This system focuses specifically on the detection of safety helmets and high-visibility reflective vests, the two most universally mandated PPE items in construction and industrial environments. It does not currently detect gloves, safety boots, eye protection, or respiratory equipment. The system processes visual data from standard RGB cameras and does not incorporate depth sensing or thermal imaging. Optimization for edge deployment on low-power hardware is noted as a key area of future scope."),
      pageBreak(),

      // ─── CHAPTER 2 ────────────────────────────────────────────
      h1("Chapter 2: Literature Review"),

      h3("2.1 Traditional PPE Monitoring Approaches"),
      p("Historically, PPE compliance in industrial settings has been enforced through a combination of mandatory training programs, posted signage, and periodic walkthroughs by designated safety officers. While these methods establish a culture of safety awareness, they suffer from fundamental limitations: they are reactive rather than proactive, they are subject to human attentional constraints, and they cannot provide the continuous monitoring required to prevent all violations. Research in occupational health consistently demonstrates that compliance rates drop significantly when workers believe they are not being observed."),

      h3("2.2 Evolution of Object Detection for Safety Applications"),
      p("The application of computer vision to workplace safety monitoring has evolved significantly over the past decade. Early approaches relied on classical image processing techniques such as color histogram analysis for high-visibility vest detection and Haar cascades for helmet recognition. While computationally inexpensive, these methods proved brittle in the face of variable lighting conditions, occlusion, and diverse worker orientations."),
      sp(120),
      p("The introduction of deep learning-based detection frameworks transformed the field. R-CNN and its successors (Fast R-CNN, Faster R-CNN) demonstrated high accuracy on PPE detection benchmarks but required two-stage processing pipelines that were too slow for real-time deployment. The YOLO architecture, first introduced by Redmon et al., solved this by reframing object detection as a single regression problem solved in one forward pass through a unified neural network, enabling real-time inference at 45+ FPS on standard hardware."),

      h3("2.3 YOLO Advances and the Ultralytics Ecosystem"),
      p("Subsequent YOLO iterations, culminating in YOLOv8 and the newer YOLO12 series from Ultralytics, introduced architectural improvements including C2f modules, anchor-free detection heads, and improved training recipes that significantly boosted mean Average Precision (mAP) on benchmark datasets such as COCO. For domain-specific tasks like PPE detection, fine-tuning pre-trained YOLO weights on labeled construction site datasets has been demonstrated to achieve mAP scores exceeding 85% for helmet and vest classification."),

      h3("2.4 The Multi-Person Association Problem"),
      p("A critical and often underaddressed challenge in PPE detection systems is the correct association of detected equipment to detected persons. In a frame containing multiple workers, a naive system might detect a helmet correctly but fail to assign it to the correct worker, leading to false compliance reports. The literature proposes several approaches including proximity-based matching, skeleton keypoint analysis using pose estimation models, and anatomical region decomposition. This project addresses the problem through an intelligent anatomical region analysis approach that defines head and torso bounding zones relative to the detected person bounding box."),
      pageBreak(),

      // ─── CHAPTER 3 ────────────────────────────────────────────
      h1("Chapter 3: Theoretical Framework"),

      h3("3.1 YOLO Architecture and Inference"),
      p("YOLO (You Only Look Once) reformulates object detection as a single regression problem. Given an input image of dimensions H x W x 3, the network divides it into an S x S grid. Each cell predicts B bounding boxes with associated confidence scores and C class probabilities simultaneously in a single forward pass."),
      sp(120),
      p("The confidence score for each bounding box is defined as:"),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 120 },
        children: [new TextRun({ text: "Confidence = Pr(Object) \u00D7 IoU(pred, truth)", italics: true, font: "Arial", size: 22 })]
      }),
      p("where IoU is the Intersection over Union between the predicted and ground truth boxes. The final class-specific confidence scores are computed as:"),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { after: 120 },
        children: [new TextRun({ text: "Pr(Class_i | Object) \u00D7 Pr(Object) \u00D7 IoU(pred, truth)", italics: true, font: "Arial", size: 22 })]
      }),
      p("Modern YOLO variants use anchor-free detection heads, eliminating the need for pre-defined anchor boxes and improving generalization to novel object aspect ratios. Non-Maximum Suppression (NMS) is applied post-inference to eliminate redundant overlapping detections, retaining only the highest-confidence box for each detected object."),

      h3("3.2 Dual-Model Detection Pipeline"),
      p("This system employs a two-stage detection strategy. In the first stage, the general-purpose YOLO person detector (yolo12n.pt) is applied to the full frame to localize all human subjects. This model is pretrained on the COCO dataset and provides high-recall person detection. In the second stage, the custom-trained PPE model (best.pt) is applied to detect helmets and vests within the context of the full frame."),
      sp(120),
      p("This dual-model approach exploits the specialization of each model. The person detector benefits from massive pretraining diversity, while the PPE detector is precisely tuned to the visual characteristics of helmets and high-visibility vests in construction environments. Post-detection, the association engine resolves which PPE items belong to which worker."),

      h3("3.3 Anatomical Region-Based PPE Association"),
      p("The core intellectual contribution of the association engine is the definition of anatomical bounding zones relative to each detected person bounding box. For a detected person with bounding box coordinates (x1, y1, x2, y2):"),
      bullet("Head Region: The top 30% of the person bounding box, corresponding to the head and neck area. Any helmet detection whose center point falls within this zone is associated with this worker."),
      bullet("Torso Region: The middle 60% of the person bounding box height, spanning the full width. Safety vest detections whose center falls within this zone are associated with this worker."),
      sp(120),
      p("This approach requires no additional pose estimation model, making it computationally efficient while remaining robust to typical worker poses encountered in construction footage."),

      h3("3.4 Temporal Smoothing for False Positive Suppression"),
      p("Single-frame detections are inherently noisy due to motion blur, partial occlusion, and lighting transients. The system maintains a detection history buffer of N consecutive frames for each tracked worker. A PPE item is only reported as absent (non-compliant) if it fails to be detected in a majority of the preceding N frames. This temporal smoothing prevents flickering compliance status and eliminates single-frame false positives from degrading the reliability of the output recording."),
      pageBreak(),

      // ─── CHAPTER 4 ────────────────────────────────────────────
      h1("Chapter 4: System Architecture & Data Pipeline"),

      h3("4.1 Dataset"),
      p("The custom PPE detection model (best.pt) was trained on the Construction Site Safety Image Dataset curated by Snehil Sanyal and published on Kaggle. The dataset comprises a substantial collection of annotated images captured from real construction sites, providing diverse coverage of worker poses, lighting conditions, and PPE configurations. Key characteristics:"),
      sp(120),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 6240],
        rows: [
          new TableRow({
            children: ["Attribute", "Detail"].map((h, i) => new TableCell({
              borders: ALL_BORDERS,
              width: { size: i === 0 ? 3120 : 6240, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              shading: { fill: "EEEEEE", type: ShadingType.CLEAR },
              children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, font: "Arial", size: 22 })] })]
            }))
          }),
          ...([
            ["Dataset Source", "Kaggle - Construction Site Safety Image Dataset (Snehil Sanyal)"],
            ["Target Classes", "Person, Helmet (Hardhat), Safety Vest (Reflective Jacket)"],
            ["Annotation Format", "YOLO format (normalized bounding box coordinates)"],
            ["Training Split", "Train / Validation / Test partitioned subsets"],
            ["Use Case", "Fine-tuning YOLO base weights for construction-site PPE detection"],
          ].map(r => new TableRow({
            children: r.map((c, i) => new TableCell({
              borders: ALL_BORDERS,
              width: { size: i === 0 ? 3120 : 6240, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: c, font: "Arial", size: 22 })] })]
            }))
          })))
        ]
      }),

      h3("4.2 System Architecture Overview"),
      p("The system follows a modular pipeline architecture with five primary components:"),
      bullet("Video Input Layer: Handles source-agnostic ingestion from webcam indices, IP camera URLs, or local video file paths using OpenCV's VideoCapture API."),
      bullet("Dual-Model Inference Layer: Applies the person detection model (yolo12n.pt) and PPE detection model (best.pt) independently to each processed frame."),
      bullet("PPE Association Engine: Maps detected helmets and vests to their corresponding workers using anatomical region analysis."),
      bullet("Temporal Smoothing Layer: Maintains per-worker detection history buffers and applies frame-consistency logic before generating compliance status."),
      bullet("Output Layer: Renders annotated frames with bounding boxes, worker IDs, and compliance status, then encodes the output stream to MP4 using OpenCV's VideoWriter."),

      h3("4.3 Project File Structure"),
      p("The repository is organized as follows:"),
      ...codeBlock([
        "Safety_Gear_Detection_Camera/",
        "  workplace_safety_monitor.py   # Main application entry point",
        "  requirements.txt               # Python dependency list",
        "  best.pt                        # Custom-trained PPE detection model",
        "  yolo12n.pt                     # General-purpose person detection model",
        "  dataset/",
        "    safety-Helmet-Reflective-Jacket/  # Training & validation data",
        "  test_videos/",
        "    veo3_construction.mp4        # Sample test footage",
        "  output/",
        "    frame_*.jpg                  # Extracted annotated frames",
        "    output.mp4                   # Final annotated video",
      ]),

      h3("4.4 Confidence Threshold Configuration"),
      p("Detection confidence thresholds for each class are configurable via command-line arguments, enabling users to tune the sensitivity-specificity tradeoff for their specific deployment environment:"),
      bullet("--conf-helmet: Minimum confidence score to accept a helmet detection (default: 0.65)"),
      bullet("--conf-vest: Minimum confidence score to accept a safety vest detection (default: 0.70)"),
      sp(120),
      p("These thresholds were determined empirically through testing on the provided construction footage. Higher values reduce false positives at the cost of potentially missing partially occluded PPE items."),
      pageBreak(),

      // ─── CHAPTER 5 ────────────────────────────────────────────
      h1("Chapter 5: Implementation & Technical Details"),

      h3("5.1 Technology Stack"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 6240],
        rows: [
          new TableRow({
            children: ["Component", "Technology / Library"].map((h, i) => new TableCell({
              borders: ALL_BORDERS,
              width: { size: i === 0 ? 3120 : 6240, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              shading: { fill: "EEEEEE", type: ShadingType.CLEAR },
              children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, font: "Arial", size: 22 })] })]
            }))
          }),
          ...([
            ["Core Language", "Python 3.x"],
            ["Computer Vision", "OpenCV (opencv-python) - video I/O, frame rendering, MP4 encoding"],
            ["Object Detection", "Ultralytics YOLO (ultralytics) - model inference"],
            ["Numerical Computing", "NumPy - array operations for region analysis"],
            ["Fallback Detection", "OpenCV HOG Pedestrian Detector (when YOLO unavailable)"],
            ["Model Weights", "best.pt (custom PPE), yolo12n.pt (person detection)"],
          ].map(r => new TableRow({
            children: r.map((c, i) => new TableCell({
              borders: ALL_BORDERS,
              width: { size: i === 0 ? 3120 : 6240, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: c, font: "Arial", size: 22 })] })]
            }))
          })))
        ]
      }),

      h3("5.2 Installation and Setup"),
      p("The system is designed for straightforward deployment. The following commands configure the environment:"),
      ...codeBlock([
        "# Clone the repository",
        "git clone https://github.com/prodbykosta/ppe-safety-detection-ai.git",
        "cd Safety_Gear_Detection_Camera",
        "",
        "# Install core dependencies",
        "pip install opencv-python numpy",
        "",
        "# Install Ultralytics (YOLO inference engine)",
        "pip install ultralytics",
        "",
        "# Verify installation",
        "python3 workplace_safety_monitor.py --help",
      ]),

      h3("5.3 Core Detection Implementation"),
      p("The main application (workplace_safety_monitor.py) implements the full pipeline. Below is a representative pseudocode illustration of the core inference and association loop:"),
      ...codeBlock([
        "from ultralytics import YOLO",
        "import cv2",
        "",
        "# Load both detection models",
        "person_model = YOLO('yolo12n.pt')",
        "ppe_model    = YOLO('best.pt')",
        "",
        "cap = cv2.VideoCapture(source)  # source: 0, URL, or file path",
        "",
        "while cap.isOpened():",
        "    ret, frame = cap.read()",
        "    if not ret: break",
        "",
        "    # Stage 1: Detect persons",
        "    persons = person_model(frame, conf=0.5)",
        "",
        "    # Stage 2: Detect PPE items",
        "    ppe_items = ppe_model(frame, conf=conf_threshold)",
        "",
        "    # Stage 3: Associate PPE to persons via anatomical regions",
        "    for person_box in persons:",
        "        head_zone  = compute_head_region(person_box)   # top 30%",
        "        torso_zone = compute_torso_region(person_box)  # mid 60%",
        "        has_helmet = any(helmet in head_zone  for helmet in ppe_items)",
        "        has_vest   = any(vest   in torso_zone for vest   in ppe_items)",
        "",
        "    # Stage 4: Apply temporal smoothing & annotate frame",
        "    annotate_and_write(frame, compliance_status)",
      ]),

      h3("5.4 Multi-Source Input Handling"),
      p("The system abstracts video source selection via the --source argument, supporting three input modes:"),
      ...codeBlock([
        "# Live webcam (default camera)",
        "python3 workplace_safety_monitor.py --ppe-weights best.pt --source 0",
        "",
        "# IP camera / RTSP stream",
        "python3 workplace_safety_monitor.py --ppe-weights best.pt --source rtsp://192.168.1.10/stream",
        "",
        "# Pre-recorded video file",
        "python3 workplace_safety_monitor.py --ppe-weights best.pt --source test_videos/veo3_construction.mp4",
      ]),

      h3("5.5 Model Evaluation"),
      p("The PPE detection model can be evaluated against a labeled validation dataset using the built-in evaluation mode:"),
      ...codeBlock([
        "python3 workplace_safety_monitor.py \\",
        "    --ppe-weights best.pt \\",
        "    --eval-root dataset/valid \\",
        "    --conf-helmet 0.65 \\",
        "    --conf-vest 0.70",
      ]),

      h3("5.6 Fallback Detection Mechanism"),
      p("A key robustness feature is the graceful fallback to OpenCV's built-in Histogram of Oriented Gradients (HOG) pedestrian detector in environments where the Ultralytics library is unavailable or model weights cannot be loaded. This ensures the system degrades gracefully rather than failing completely, maintaining at least person detection capability for environments with constrained dependencies."),
      pageBreak(),

      // ─── CHAPTER 6 ────────────────────────────────────────────
      h1("Chapter 6: Results and Discussion"),

      h3("6.1 Detection Performance"),
      p("The system was tested on the provided sample construction footage (veo3_construction.mp4) featuring multiple workers in various poses and lighting conditions. The dual-model architecture demonstrated strong person detection recall, with the yolo12n.pt model successfully identifying workers across varying distances from the camera."),
      sp(120),
      p("PPE detection accuracy was evaluated at the configured confidence thresholds of 0.65 for helmets and 0.70 for vests. Key qualitative findings:"),
      bullet("Helmet Detection: High precision was observed for workers facing the camera directly. Performance degraded moderately for workers seen from behind, which is consistent with known YOLO limitations on rear-facing occluded objects."),
      bullet("Vest Detection: The high-visibility nature of reflective safety vests provided a strong visual signal even under variable lighting, contributing to reliable vest detection across most test scenarios."),
      bullet("Association Accuracy: The anatomical region approach correctly attributed PPE items to workers in the majority of test frames, with occasional mis-associations occurring only in heavily crowded scenes with significant overlap between worker bounding boxes."),

      h3("6.2 Temporal Smoothing Impact"),
      p("Temporal smoothing provided a measurable reduction in flickering compliance status. Without smoothing, compliance labels toggled rapidly between frames due to detection variability. With the frame-consistency buffer active, the output video displayed stable, readable compliance annotations that accurately reflected the sustained PPE status of each worker."),

      h3("6.3 Sequential Worker ID Assignment"),
      p("Workers detected within the video stream are assigned sequential numeric IDs upon their first appearance. These IDs persist across frames through tracking logic, enabling per-worker compliance history to be maintained across the duration of a recording. This ID stability is critical for generating meaningful per-worker audit records."),

      h3("6.4 Output Video Quality"),
      p("The annotated output MP4 is encoded at the source video resolution, preserving image quality for review purposes. Bounding boxes are rendered in distinct colors per compliance category (e.g., green for compliant, red for non-compliant), and worker IDs are displayed as persistent overlays. The output is stored in the output/ directory and is suitable for direct playback in standard video players."),
      pageBreak(),

      // ─── CHAPTER 7 ────────────────────────────────────────────
      h1("Chapter 7: Societal Impact and Commercial Value"),

      h3("7.1 Workplace Safety Impact"),
      p("Workplace injuries attributable to PPE non-compliance impose enormous human and economic costs on society. Beyond the immediate physical harm to individual workers, violations result in regulatory fines, increased insurance premiums, construction delays, and long-term litigation. An automated, continuously operating monitoring system directly addresses the enforcement gap created by infrequent manual inspections, enabling site managers to identify and respond to violations in near real-time."),
      sp(120),
      p("The system's ability to generate annotated video recordings also creates an objective evidence record. In the event of an incident, this footage can provide documentation of the safety conditions at the time, supporting both insurance claims and regulatory reporting requirements."),

      h3("7.2 Commercial Viability"),
      p("The commercial case for AI-powered PPE monitoring is strong. The system is built entirely on open-source foundations (OpenCV, Ultralytics, Python), making the marginal cost of deployment minimal compared to proprietary safety monitoring hardware. The support for IP cameras means the system can be integrated into existing CCTV infrastructure without requiring new camera hardware."),
      sp(120),
      p("Potential commercial deployment models include a Software-as-a-Service (SaaS) safety compliance platform, an on-premises appliance sold directly to large construction or manufacturing enterprises, or integration as a module within existing Building Management System (BMS) or ERP safety modules."),

      h3("7.3 Environmental and Operational Sustainability"),
      p("By providing data-driven insight into compliance patterns across shifts and locations, the system enables safety managers to target training interventions precisely where and when they are most needed. This reduces the operational waste associated with blanket retraining programs and focuses human safety effort on the highest-risk scenarios identified by the AI monitor."),
      pageBreak(),

      // ─── CHAPTER 8 ────────────────────────────────────────────
      h1("Chapter 8: Conclusion and Future Scope"),

      h3("8.1 Conclusion"),
      p("This project has successfully demonstrated a complete, deployable AI-powered PPE Safety Detection System for real-time workplace safety monitoring. The system achieves its core objectives: detecting persons and their PPE compliance status using a dual YOLO model architecture, correctly associating equipment to workers through anatomical region analysis, suppressing false positives via temporal smoothing, and producing annotated MP4 output suitable for audit and training purposes."),
      sp(120),
      p("The architecture is practically deployable with standard consumer hardware and integrates seamlessly with existing camera infrastructure. The use of open-source libraries and a custom-trained model on a publicly available dataset demonstrates a reproducible and extensible approach to industrial AI applications."),

      h3("8.2 Future Scope"),
      p("Several directions for system enhancement are identified:"),
      bullet("Extended PPE Classes: Adding detection support for gloves, safety footwear, eye protection, and respiratory masks to provide comprehensive compliance monitoring."),
      bullet("Edge Deployment: Optimizing the model through quantization and pruning for deployment on low-power edge devices such as NVIDIA Jetson Nano or Raspberry Pi with Coral TPU accelerators, eliminating the need for cloud connectivity."),
      bullet("Alert System Integration: Implementing real-time alert triggering via MQTT messages or REST API calls to safety officer mobile applications when violations are detected for a sustained duration."),
      bullet("Analytics Dashboard: Building a web-based dashboard that aggregates per-worker, per-zone compliance statistics over time, enabling trend analysis and predictive safety management."),
      bullet("Multi-Camera Synchronization: Extending the architecture to handle synchronized feeds from multiple cameras across large facilities, with cross-camera worker re-identification."),
      pageBreak(),

      // ─── REFERENCES ───────────────────────────────────────────
      h2("References"),
      p("[1] Jocher, G. et al. (2023). Ultralytics YOLO. GitHub. https://github.com/ultralytics/ultralytics"),
      p("[2] Sanyal, S. (2022). Construction Site Safety Image Dataset. Kaggle. https://www.kaggle.com/datasets/snehilsanyal/construction-site-safety-image-dataset-roboflow"),
      p("[3] Redmon, J., Divvala, S., Girshick, R., & Farhadi, A. (2016). You Only Look Once: Unified, Real-Time Object Detection. CVPR 2016."),
      p("[4] Bradski, G. (2000). The OpenCV Library. Dr. Dobb's Journal of Software Tools."),
      p("[5] prodbykosta. (2025). ppe-safety-detection-ai. GitHub. https://github.com/prodbykosta/ppe-safety-detection-ai"),
      p("[6] Bochkovskiy, A., Wang, C.Y., & Liao, H.Y.M. (2020). YOLOv4: Optimal Speed and Accuracy of Object Detection. arXiv:2004.10934."),
      p("[7] Lundberg, S.M. & Lee, S.I. (2017). A Unified Approach to Interpreting Model Predictions. NeurIPS 2017."),
      pageBreak(),

      // ─── APPENDIX ─────────────────────────────────────────────
      h2("Appendix: LLM Usage Report (Annexure 3b)"),
      p("This appendix documents the use of Large Language Models (LLMs) and AI tools during the development of the PPE Safety Detection project, in accordance with institutional disclosure requirements."),

      h3("Section A: Project Overview"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 6240],
        rows: [
          ...([
            ["Project Title", "AI-Powered PPE Safety Detection System"],
            ["Registration No.", "[Your Registration No.]"],
            ["Faculty Guide", "Dr. Jimmy Singla"],
            ["Date", "[Submission Date]"],
          ].map(r => new TableRow({
            children: r.map((c, i) => new TableCell({
              borders: ALL_BORDERS,
              width: { size: i === 0 ? 3120 : 6240, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: c, bold: i === 0, font: "Arial", size: 22 })] })]
            }))
          })))
        ]
      }),

      h3("Section B: AI Tools Used"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 2340, 4680],
        rows: [
          new TableRow({
            children: ["AI Tool / LLM", "Purpose", "Usage Description"].map((h, i) => new TableCell({
              borders: ALL_BORDERS,
              width: { size: [2340, 2340, 4680][i], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              shading: { fill: "EEEEEE", type: ShadingType.CLEAR },
              children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, font: "Arial", size: 22 })] })]
            }))
          }),
          ...([
            ["GitHub Copilot / ChatGPT", "Code Assistance", "Used to assist in debugging the YOLO inference loop and refining the anatomical region association logic."],
            ["Claude (Anthropic)", "Documentation", "Assisted in structuring and drafting the project report, ensuring technical accuracy of architectural descriptions."],
            ["Gemini (Google)", "Research Support", "Used for literature review summaries and understanding recent YOLO architecture improvements."],
          ].map(r => new TableRow({
            children: r.map((c, i) => new TableCell({
              borders: ALL_BORDERS,
              width: { size: [2340, 2340, 4680][i], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: c, font: "Arial", size: 22 })] })]
            }))
          })))
        ]
      }),

      h3("Section C: LLM Interaction Details"),
      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun({ text: "System Prompting Strategy", font: "Arial", size: 26 })]
      }),
      p("For code generation tasks, AI tools were prompted with explicit context about the project's dual-model architecture, the specific YOLO version (yolo12n), and the PPE classes being detected. Prompts explicitly requested that outputs adhere to OpenCV's VideoCapture and VideoWriter API conventions."),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun({ text: "Prompt Examples Used", font: "Arial", size: 26 })]
      }),
      ...codeBlock([
        "Prompt 1: 'Given a YOLO bounding box (x1, y1, x2, y2) for a detected person,",
        "write a Python function that computes the head region (top 30%) and torso region",
        "(middle 60%) as separate bounding boxes for PPE association.'",
        "",
        "Prompt 2: 'How do I implement temporal smoothing for object detection compliance",
        "status over N frames using a deque in Python?'",
      ]),

      h3("Section D: Model Configuration & Behavior"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 3120, 3900],
        rows: [
          new TableRow({
            children: ["Parameter", "Value", "Applicability"].map((h, i) => new TableCell({
              borders: ALL_BORDERS,
              width: { size: [2340, 3120, 3900][i], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              shading: { fill: "EEEEEE", type: ShadingType.CLEAR },
              children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, font: "Arial", size: 22 })] })]
            }))
          }),
          ...([
            ["YOLO Confidence (Helmet)", "0.65", "PPE Detection Model"],
            ["YOLO Confidence (Vest)", "0.70", "PPE Detection Model"],
            ["YOLO Confidence (Person)", "0.50", "Person Detection Model"],
            ["Temporal Buffer Length", "N frames (configurable)", "Smoothing Layer"],
            ["LLM Temperature (doc assist)", "0.2", "Report Generation Assistance"],
          ].map(r => new TableRow({
            children: r.map((c, i) => new TableCell({
              borders: ALL_BORDERS,
              width: { size: [2340, 3120, 3900][i], type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: c, font: "Arial", size: 22 })] })]
            }))
          })))
        ]
      }),

      h3("Section E: Technology Stack Summary"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 6240],
        rows: [
          ...([
            ["Computer Vision", "OpenCV (opencv-python)"],
            ["Object Detection", "Ultralytics YOLO (ultralytics)"],
            ["Person Model", "yolo12n.pt (pretrained on COCO)"],
            ["PPE Model", "best.pt (fine-tuned on Construction Safety Dataset)"],
            ["Numerical Computing", "NumPy"],
            ["Language", "Python 3.x"],
            ["Output Format", "Annotated MP4 video"],
            ["Hosting / Deployment", "Local / Edge device (no cloud dependency required)"],
          ].map(r => new TableRow({
            children: r.map((c, i) => new TableCell({
              borders: ALL_BORDERS,
              width: { size: i === 0 ? 3120 : 6240, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: c, bold: i === 0, font: "Arial", size: 22 })] })]
            }))
          })))
        ]
      }),

      h3("Section F: GitHub Repository Link"),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({ text: "Repository URL: ", bold: true, font: "Arial", size: 22 }),
          new ExternalHyperlink({
            children: [new TextRun({ text: "https://github.com/prodbykosta/ppe-safety-detection-ai", style: "Hyperlink", font: "Arial", size: 22 })],
            link: "https://github.com/prodbykosta/ppe-safety-detection-ai"
          })
        ]
      }),
      sp(120),
      p("Note: All code, model weights, test videos, and output samples referenced in this report are available in the above repository. The custom-trained model (best.pt) and person detector (yolo12n.pt) are included directly in the repository for ease of deployment."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("./INT428_PPE_Safety_Detection_Formatted.docx", buffer);
  console.log("Done!");
}).catch(e => console.error(e));