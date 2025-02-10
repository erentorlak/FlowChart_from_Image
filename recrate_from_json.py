#!/usr/bin/env python

#%%
from ctypes.wintypes import RGB
import json
import win32com.client
# -------------------------------
# Global Shape and OCR Mappings
# -------------------------------
SHAPE_DRAW_MAP = {
    4: 4,    # Decision (Diamond)
    5: None, # Arrow tip (ignored)
    6: 67,    # Output Box
    7: 1,    # Rectangle (Process Box)
    8: 2,   # Trapezoid
    9: 9     # Ellipse (Start/End)
}

def load_chart_data(json_path):
    """Load chart data from JSON file."""
    with open(json_path, "r") as f:
        return json.load(f)

def initialize_word():
    """Initialize Word application and create a new document."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Add()
    return word, doc

def create_shapes(doc, nodes):
    """
    Create shapes in the document based on node data.
    If an OCR text is present in the node, it is used as the shape text.
    Otherwise, the node's "name" is used.
    """
    node_objects = {}
    shapes = doc.Shapes
    
    for node in nodes:
        shape_id = node["id"]
        class_id = node["class_id"]
        x1, y1, width, height = node["bbox"]
        # Use the same drawing mapping as originally.
        shape_type = SHAPE_DRAW_MAP.get(class_id, 1)
        
        shp = shapes.AddShape(shape_type, x1, y1, width, height)
        # Set fill color based on class.
        if class_id == 4:
            shp.Fill.ForeColor.RGB = RGB(0, 255, 0)
        elif class_id == 6:
            shp.Fill.ForeColor.RGB = RGB(0, 0, 255)
        elif class_id == 7:
            shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
        elif class_id == 8:
            shp.Fill.ForeColor.RGB = RGB(255, 255, 0)
        elif class_id == 9:
            shp.Fill.ForeColor.RGB = RGB(0, 255, 255)
        # Use OCR text if available; otherwise use the node name.
        if node.get("ocr_text"):
            shp.TextFrame.TextRange.Text = node["ocr_text"]
        elif node.get("name"):
            shp.TextFrame.TextRange.Text = node["name"]
        else:
            shp.TextFrame.TextRange.Text = ""
        node_objects[shape_id] = shp
    
    return node_objects

def create_arrows(doc, arrows):
    """Create arrows connecting the shapes using the provided arrow endpoints."""
    shapes = doc.Shapes
    
    for arrow in arrows:
        # Arrow endpoints are expected to be stored as lists [x, y]
        tail_pt = arrow["tail"]["point"]
        tip_pt = arrow["tip"]["point"]
        line = shapes.AddLine(tail_pt[0], tail_pt[1], tip_pt[0], tip_pt[1])
        line.Line.EndArrowheadStyle = 3

def save_and_close(doc, word, doc_path):
    """Save the document and close Word."""
    doc.SaveAs(doc_path)
    doc.Close()
    word.Quit()

def main():
    """Main function to recreate the chart from JSON."""
    json_path = "C:\\Users\\Eren.Torlak\\OneDrive - Logo\\Desktop\\net5.json"
    doc_path = "C:\\Users\\Eren.Torlak\\OneDrive - Logo\\Desktop\\net5_recreated.docx"
    
    # Load chart data from JSON.
    chart = load_chart_data(json_path)
    
    # Initialize Word.
    word, doc = initialize_word()
    
    # Create shapes and then arrows.
    node_objects = create_shapes(doc, chart["nodes"])
    create_arrows(doc, chart.get("arrows", []))
    
    # Save and close the document.
    save_and_close(doc, word, doc_path)
    print(f"Chart reproduced and saved at: {doc_path}")

if __name__ == "__main__":
    main()

# %%
