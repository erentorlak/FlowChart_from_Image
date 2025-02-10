#!/usr/bin/env python
#%%
"""
recreate_from_nxnmatrix.py

This script loads an NXN matrix JSON file (with keys "matrix" and "shapes"),
recreates the shapes in a new Word document using the original drawing parameters,
and draws arrows between shapes according to the matrix connections.

The JSON file is expected to have been produced by your earlier flowchart generation,
where "shapes" is a list of shape mapping dictionaries (each with "id", "class_id",
"name", "bbox", "center", and optionally "ocr_text") and "matrix" is a 2D list where
matrix[i][j]==1 indicates an edge from shape i to shape j.
"""

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

def load_matrix_data(json_path):
    """Load NXN matrix and shape mapping data from JSON file."""
    with open(json_path, "r") as f:
        return json.load(f)

def initialize_word():
    """Initialize Word application and create a new document."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Add()
    return word, doc

def create_shapes(doc, shapes):
    """
    Create shapes in the document based on shape mapping data.
    
    Each shape mapping should include:
      - "id": the shape's unique identifier.
      - "class_id": numeric ID (used to choose the drawn shape type).
      - "name": human-readable label (or OCR text if available).
      - "bbox": [x1, y1, width, height].
      - "center": [cx, cy] (center coordinates).
    
    Returns a dictionary mapping shape id to the created Word shape object.
    """
    node_objects = {}
    shapes_collection = doc.Shapes
    for shape in shapes:
        shape_id = shape["id"]
        class_id = shape["class_id"]
        x1, y1, width, height = shape["bbox"]
        shape_type = SHAPE_DRAW_MAP.get(class_id, 1)
        shp = shapes_collection.AddShape(shape_type, x1, y1, width, height)
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
        # Use OCR text if present, otherwise the provided name.
        if shape.get("ocr_text"):
            shp.TextFrame.TextRange.Text = shape["ocr_text"]
        elif shape.get("name"):
            shp.TextFrame.TextRange.Text = shape["name"]
        else:
            shp.TextFrame.TextRange.Text = ""
        node_objects[shape_id] = shp
    return node_objects

def compute_arrow_endpoints(source, target):
    """
    Compute arrow endpoints between two shapes using a simple heuristic.
    
    Parameters:
      source: A shape mapping dictionary for the source.
      target: A shape mapping dictionary for the target.
      
    Returns:
      A tuple (source_pt, target_pt) where each is a (x, y) tuple.
      
    Heuristic:
      - If source center is above target center, use the bottom edge of source and top edge of target.
      - If source center is below target center, use the top edge of source and bottom edge of target.
      - Otherwise, if source is left of target, use the right edge of source and left edge of target.
      - Else, use left edge of source and right edge of target.
    """
    sx, sy, sw, sh = source["bbox"]
    scx, scy = source["center"]
    tx, ty, tw, th = target["bbox"]
    tcx, tcy = target["center"]
    
    if scy < tcy:
        source_pt = (scx, sy + sh)   # bottom center of source
        target_pt = (tcx, ty)        # top center of target
    elif scy > tcy:
        source_pt = (scx, sy)        # top center of source
        target_pt = (tcx, ty + th)   # bottom center of target
    elif scx < tcx:
        source_pt = (sx + sw, scy)   # right center of source
        target_pt = (tx, tcy)        # left center of target
    else:
        source_pt = (sx, scy)        # left center of source
        target_pt = (tx + tw, tcy)   # right center of target
    return source_pt, target_pt

def create_arrows_from_matrix(doc, matrix, shape_mapping):
    """
    Create arrows in the Word document based on the NXN matrix.
    
    For each cell matrix[i][j] == 1, an arrow is drawn from shape with id i to shape with id j.
    
    Parameters:
      doc: Word document.
      matrix: NXN matrix (list of lists) representing connections.
      shape_mapping: List of shape mapping dictionaries.
    """
    shapes_collection = doc.Shapes
    # Build a dictionary for easy lookup of shapes by id.
    shape_dict = {shape["id"]: shape for shape in shape_mapping}
    
    n = len(matrix)
    for i in range(n):
        for j in range(n):
            if matrix[i][j] == 1:
                source = shape_dict.get(i)
                target = shape_dict.get(j)
                if source and target:
                    source_pt, target_pt = compute_arrow_endpoints(source, target)
                    line = shapes_collection.AddLine(source_pt[0], source_pt[1],
                                                     target_pt[0], target_pt[1])
                    line.Line.EndArrowheadStyle = 3

def save_and_close(doc, word, doc_path):
    """Save the document and close Word."""
    doc.SaveAs(doc_path)
    doc.Close()
    word.Quit()

def main():
    """Main function to recreate the chart from the NXN matrix JSON data."""
    json_path = "C:\\Users\\Eren.Torlak\\OneDrive - Logo\\Desktop\\net_matrix5.json"
    doc_path = "C:\\Users\\Eren.Torlak\\OneDrive - Logo\\Desktop\\net_recreated_matrix5.docx"
    
    # Load the NXN matrix data (expects keys "matrix" and "shapes").
    with open(json_path, "r") as f:
        data = json.load(f)
    matrix = data["matrix"]
    shape_mapping = data["shapes"]
    
    # Initialize Word.
    word, doc = initialize_word()
    
    # Create shapes.
    create_shapes(doc, shape_mapping)
    
    # Create arrows based on the NXN matrix.
    create_arrows_from_matrix(doc, matrix, shape_mapping)
    
    # Save and close the document.
    save_and_close(doc, word, doc_path)
    print(f"Chart reproduced from NXN matrix and saved at: {doc_path}")

if __name__ == "__main__":
    main()

# %%
