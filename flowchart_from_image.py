#%%

# %pip install ultralytics python-docx pywin32 opencv-python transformers pillow openai

import math, json, cv2, base64, io, re
from collections import defaultdict, deque
from ultralytics import YOLO
import win32com.client
from ctypes.wintypes import RGB
from PIL import Image
import openai


# -------------------------------
# Global Shape and OCR Mappings
# -------------------------------
SHAPE_MAP = {
    4: 4,    # Decision (Diamond)
    5: None, # Arrow tip (ignored)
    6: 67,    # output
    7: 1,    # Rectangle (Process Box)
    8: 2,       # parelelkenar 
    9: 9     # Ellipse (Start/End)
}
SHAPE_NAME_MAP = {
    4: "Decision",
    5: None,
    6: "Output",
    7: "Process",
    8: "Scan (input)",
    9: "Start/End"
}

# -------------------------------
# Global Connection Node Counter
# -------------------------------
connection_id_counter = 0
def get_next_connection_id():
    """Return the next unique connection node ID."""
    global connection_id_counter
    cid = connection_id_counter
    connection_id_counter += 1
    return cid

# -------------------------------
# OCR Functionality Using GPT API
# -------------------------------
def add_gpt_ocr_to_shapes(shape_positions, image_path, padding=5):
    """
    For each shape, extract handwritten OCR text from its image region using GPT API.
    The prompt instructs the model to "Return only the OCR output" (no extra commentary).
    
    Parameters:
      shape_positions: List of shapes (each is a tuple with bounding box info).
      image_path: Path to the source image.
      padding: Extra pixels to add around the ROI.
    
    Returns:
      A dictionary mapping shape_id to the extracted OCR text.
    """
    ocr_dict = {}
    # Open the source image using PIL.
    image = Image.open(image_path).convert("RGB")
    for shape in shape_positions:
        shape_id, _, x1, y1, width, height, _, _, _, _ = shape
        # Compute padded ROI coordinates.
        x1_p = max(0, x1 - padding)
        y1_p = max(0, y1 - padding)
        x2_p = x1 + width + padding
        y2_p = y1 + height + padding
        roi = image.crop((x1_p, y1_p, x2_p, y2_p))
        # Encode ROI to base64.
        buffer = io.BytesIO()
        roi.save(buffer, format="JPEG")
        base64_image = base64.b64encode(buffer.getvalue()).decode("utf-8")
        # Build the prompt message.
        client = openai.OpenAI(api_key="YOURAPİKEY")

        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "Return only the OCR output:"},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                ]
            }
        ]
        # Call the GPT API.
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            temperature=0.0,
            max_tokens=256,
            top_p=1
        )
        text = response.choices[0].message.content
        ocr_dict[shape_id] = text.strip()
    return ocr_dict

def icerigi_al(metin):
    """
    Extract text inside quotes from the provided string.
    If no quotes are found, return the original string.
    """
    eslesen = re.search(r'"(.*?)"', metin)
    return eslesen.group(1) if eslesen else metin

# -------------------------------
# Drawing and Connection Functions
# -------------------------------
def draw_shapes(results, shapes_obj):
    """
    Draw non-arrow shapes in the Word document using YOLO results.
    Each shape is stored as:
       (shape_id, class_id, x1, y1, width, height, cx, cy, edge_centers, shape_obj)
    
    Parameters:
       results: YOLO detection results.
       shapes_obj: Word document Shapes collection.
    
    Returns:
       List of shape information.
    """
    shape_positions = []
    for result in results:
        for box in result.boxes:
            x1, y1, x2, y2 = map(int, box.xyxy[0])
            class_id = int(box.cls[0])
            width = x2 - x1
            height = y2 - y1
            if class_id not in [0, 1, 2, 3] and class_id != 5:
                shape_type = SHAPE_MAP.get(class_id, None)
                if shape_type is None:
                    continue
                shape_obj = shapes_obj.AddShape(shape_type, x1, y1, width, height)
                if class_id == 4:
                    shape_obj.Fill.ForeColor.RGB = RGB(0, 255, 0)
                elif class_id == 6:
                    shape_obj.Fill.ForeColor.RGB = RGB(0, 0, 255)
                elif class_id == 7:
                    shape_obj.Fill.ForeColor.RGB = RGB(255, 0, 0)
                elif class_id == 8:
                    shape_obj.Fill.ForeColor.RGB = RGB(255, 255, 0)
                elif class_id == 9:
                    shape_obj.Fill.ForeColor.RGB = RGB(0, 255, 255)
                class_name = SHAPE_NAME_MAP.get(class_id, f"Class {class_id}")
                shape_obj.TextFrame.TextRange.Text = class_name
                cx, cy = x1 + width/2, y1 + height/2
                edge_centers = {
                    'top': (cx, y1),
                    'bottom': (cx, y1 + height),
                    'left': (x1, cy),
                    'right': (x1 + width, cy)
                }
                shape_id = len(shape_positions)
                shape_positions.append((shape_id, class_id, x1, y1, width, height, cx, cy, edge_centers, shape_obj))
    return shape_positions

def build_connections_pool(shape_positions):
    """
    Build a pool of connection nodes from each shape's edge centers.
    
    Parameters:
       shape_positions: List of shapes.
    
    Returns:
       List of connection nodes.
    """
    connections_pool = []
    for shape in shape_positions:
        shape_id, class_id, x1, y1, width, height, cx, cy, edge_centers, _ = shape
        for side, pt in edge_centers.items():
            conn = {
                'id': get_next_connection_id(),
                'point': pt,
                'type': 'shape',
                'shape_id': shape_id,
                'side': side
            }
            connections_pool.append(conn)
    return connections_pool

def find_nearest_connection(candidate, pool, threshold=50):
    """
    For a candidate point, search the pool for a connection node within threshold.
    If none is found, create a new connection node (of type 'arrow') and add it.
    
    Parameters:
       candidate: (x, y) tuple.
       pool: List of connection nodes.
       threshold: Distance threshold.
    
    Returns:
       A connection node dictionary.
    """
    best = None
    best_dist = threshold
    for conn in pool:
        d = math.hypot(candidate[0] - conn['point'][0], candidate[1] - conn['point'][1])
        if d < best_dist:
            best_dist = d
            best = conn
    if best is None:
        best = {
            'id': get_next_connection_id(),
            'point': candidate,
            'type': 'arrow',
            'shape_id': None,
            'side': None
        }
        pool.append(best)
    return best

def process_arrow_detections(results, connections_pool, shapes_obj):
    """
    Process arrow detections: compute candidate endpoints, snap them,
    and draw the arrow in Word.
    
    Parameters:
       results: YOLO detection results.
       connections_pool: Pool of connection nodes.
       shapes_obj: Word document Shapes collection.
    
    Returns:
       List of arrow connection details.
    """
    arrow_detections = []
    for result in results:
        for box in result.boxes:
            x1, y1, x2, y2 = map(int, box.xyxy[0])
            class_id = int(box.cls[0])
            if class_id in [0, 1, 2, 3]:
                arrow_detections.append((class_id, x1, y1, x2, y2))
    
    arrow_connections = []
    for arrow in arrow_detections:
        class_id, ax1, ay1, ax2, ay2 = arrow
        if class_id == 0:
            tail_candidate = ((ax1 + ax2) / 2, ay1)
            tip_candidate  = ((ax1 + ax2) / 2, ay2)
        elif class_id == 3:
            tail_candidate = ((ax1 + ax2) / 2, ay2)
            tip_candidate  = ((ax1 + ax2) / 2, ay1)
        elif class_id == 1:
            tail_candidate = (ax2, (ay1 + ay2) / 2)
            tip_candidate  = (ax1, (ay1 + ay2) / 2)
        elif class_id == 2:
            tail_candidate = (ax1, (ay1 + ay2) / 2)
            tip_candidate  = (ax2, (ay1 + ay2) / 2)
        else:
            continue
        tail_conn = find_nearest_connection(tail_candidate, connections_pool, threshold=50)
        tip_conn = find_nearest_connection(tip_candidate, connections_pool, threshold=50)
        line = shapes_obj.AddLine(tail_conn['point'][0], tail_conn['point'][1],
                                  tip_conn['point'][0], tip_conn['point'][1])
        line.Line.EndArrowheadStyle = 3
        arrow_connections.append({
            'class_id': class_id,
            'bbox': [ax1, ay1, ax2, ay2],
            'tail_candidate': list(tail_candidate),
            'tip_candidate': list(tip_candidate),
            'tail': tail_conn,
            'tip': tip_conn
        })
    return arrow_connections

def get_node_by_id(node_id, pool):
    """
    Retrieve a connection node from the pool by its ID.
    
    Parameters:
       node_id: Connection node ID.
       pool: List of connection nodes.
    
    Returns:
       The connection node dictionary or None.
    """
    for conn in pool:
        if conn['id'] == node_id:
            return conn
    return None

def collapse_arrow_chains(connections_pool, arrow_connections):
    """
    Collapse chains of arrow connections to produce final shape-to-shape edges.
    
    Parameters:
       connections_pool: List of connection nodes.
       arrow_connections: List of arrow connection details.
    
    Returns:
       List of final edges as tuples (from_shape_id, to_shape_id).
    """
    graph = defaultdict(list)
    for arrow in arrow_connections:
        tail = arrow['tail']
        tip = arrow['tip']
        graph[tail['id']].append(tip['id'])
    
    final_edges = set()
    for conn in connections_pool:
        if conn['type'] == 'shape':
            start_id = conn['id']
            start_shape = conn['shape_id']
            visited = set()
            queue = deque([start_id])
            while queue:
                current_id = queue.popleft()
                if current_id in visited:
                    continue
                visited.add(current_id)
                if current_id != start_id:
                    node = get_node_by_id(current_id, connections_pool)
                    if node and node['type'] == 'shape':
                        final_edges.add((start_shape, node['shape_id']))
                for neighbor in graph[current_id]:
                    if neighbor not in visited:
                        queue.append(neighbor)
    final_edges = [edge for edge in final_edges if edge[0] != edge[1]]
    return final_edges

def build_chart_json(shape_positions, final_edges, arrow_connections, ocr_dict):
    """
    Build the JSON representation of the chart including OCR text.
    
    Parameters:
       shape_positions: List of shapes.
       final_edges: Final shape-to-shape edges.
       arrow_connections: List of arrow connection details.
       ocr_dict: Mapping of shape_id to OCR text.
    
    Returns:
       Dictionary representing the chart.
    """
    chart_json = {'nodes': [], 'edges': []}
    for shape in shape_positions:
        shape_id, class_id, x1, y1, width, height, cx, cy, edge_centers, _ = shape
        chart_json['nodes'].append({
            'id': shape_id,
            'class_id': class_id,
            'name': SHAPE_NAME_MAP.get(class_id),
            'bbox': [x1, y1, width, height],
            'center': [cx, cy],
            'edge_centers': {k: list(v) for k, v in edge_centers.items()},
            'ocr_text': ocr_dict.get(shape_id, "")
        })
    for edge in final_edges:
        chart_json['edges'].append({'from': edge[0], 'to': edge[1]})
    chart_json["SHAPE_MAP"] = {str(k): v for k, v in SHAPE_NAME_MAP.items()}
    chart_json["arrows"] = []
    for arrow in arrow_connections:
        chart_json["arrows"].append({
            'class_id': arrow['class_id'],
            'bbox': arrow['bbox'],
            'tail_candidate': arrow['tail_candidate'],
            'tip_candidate': arrow['tip_candidate'],
            'tail': {
                'id': arrow['tail']['id'],
                'point': list(arrow['tail']['point']),
                'type': arrow['tail']['type'],
                'shape_id': arrow['tail']['shape_id'],
                'side': arrow['tail']['side']
            },
            'tip': {
                'id': arrow['tip']['id'],
                'point': list(arrow['tip']['point']),
                'type': arrow['tip']['type'],
                'shape_id': arrow['tip']['shape_id'],
                'side': arrow['tip']['side']
            }
        })
    return chart_json

def build_nxn_matrix(shape_positions, final_edges):
    """
    Build an NXN matrix and a mapping of shape details.
    
    Parameters:
       shape_positions: List of shapes.
       final_edges: List of final edges (from_shape_id, to_shape_id).
    
    Returns:
       Tuple (matrix, shape_mapping)
    """
    n = len(shape_positions)
    matrix = [[0] * n for _ in range(n)]
    for edge in final_edges:
        i, j = edge
        matrix[i][j] = 1
    shape_mapping = []
    for shape in shape_positions:
        shape_id, class_id, x1, y1, width, height, cx, cy, _, _ = shape
        shape_mapping.append({
            "id": shape_id,
            "class_id": class_id,
            "name": SHAPE_NAME_MAP.get(class_id, f"Class {class_id}"),
            "bbox": [x1, y1, width, height],
            "center": [cx, cy]
        })
    return matrix, shape_mapping

# -------------------------------
# Main Processing Function
# -------------------------------
def main():
    """
    Main function to:
      1. Load YOLO results.
      2. Draw shapes in Word.
      3. Build connection pool.
      4. Process arrow detections.
      5. Perform OCR using GPT API and update shape texts.
      6. Collapse arrow chains.
      7. Build JSON and NXN matrix.
      8. Save outputs.
    """
    # Load YOLO detections.
    model = YOLO("best.pt")
    image_path = "test\\images\\net.jpg"  
    results = model(image_path)
    
    # Start Word and create a new document.
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Add()
    shapes_obj = doc.Shapes
    
    # STEP 1: Draw shapes and record their info (store shape object too).
    shape_positions = draw_shapes(results, shapes_obj)
    
    # STEP 2: Build connection pool from shape edge centers.
    connections_pool = build_connections_pool(shape_positions)
    
    # STEP 3: Process arrow detections and draw arrows.
    arrow_connections = process_arrow_detections(results, connections_pool, shapes_obj)
    
    # STEP 4: Perform OCR on each shape region using GPT API.
    ocr_dict = add_gpt_ocr_to_shapes(shape_positions, image_path, padding=5)
    
    # Optional: extract text inside quotes if needed.
    def icerigi_al(metin):
        eslesen = re.search(r'"(.*?)"', metin)
        return eslesen.group(1) if eslesen else metin

    # STEP 5: Update each Word shape with OCR text (only the OCR output).
    for shape in shape_positions:
        shape_id, class_id, x1, y1, width, height, cx, cy, edge_centers, shape_obj = shape
        ocr_text = ocr_dict.get(shape_id, "")
        ocr_text = icerigi_al(ocr_text)
        shape_obj.TextFrame.TextRange.Text = str(ocr_text)
    
    # STEP 6: Save Word document.
    doc_path = "C:\\Users\\Eren.Torlak\\OneDrive - Logo\\Desktop\\net5.docx"
    doc.SaveAs(doc_path)
    doc.Close()
    word.Quit()
    print(f"Flowchart saved at: {doc_path}")
    
    # STEP 7: Collapse arrow chains to generate final edges.
    final_edges = collapse_arrow_chains(connections_pool, arrow_connections)
    
    # STEP 8: Build JSON representation of the chart.
    chart_json = build_chart_json(shape_positions, final_edges, arrow_connections, ocr_dict)
    json_output = json.dumps(chart_json, indent=4)
    print("JSON Representation of Chart:")
    print(json_output)
    with open("C:\\Users\\Eren.Torlak\\OneDrive - Logo\\Desktop\\net5.json", "w") as f:
        f.write(json_output)
    
    # STEP 9: Build NXN matrix and shape mapping JSON.
    matrix, shape_mapping = build_nxn_matrix(shape_positions, final_edges)
    print("NXN Matrix:")
    for row in matrix:
        print(row)
    matrix_info = {"matrix": matrix, "shapes": shape_mapping}
    matrix_json_output = json.dumps(matrix_info, indent=4)
    print("NXN Matrix with Shape Mapping JSON:")
    print(matrix_json_output)
    with open("C:\\Users\\Eren.Torlak\\OneDrive - Logo\\Desktop\\net_matrix5.json", "w") as f:
        f.write(matrix_json_output)

if __name__ == "__main__":
    main()
#%%
