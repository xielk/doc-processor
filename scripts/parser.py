import json
import re
import sys
import os
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

# -------------------------------------------------------
# Iteration Helper
# -------------------------------------------------------
def iter_block_items(parent):
    """
    Iterate over Paragraph and Table items in document order.
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    elif hasattr(parent, 'element') and hasattr(parent.element, 'body'):
        parent_elm = parent.element.body
    else:
        raise ValueError(f"Unsupported parent type: {type(parent)}")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# -------------------------------------------------------
# Core Parser
# -------------------------------------------------------
def parse_docx(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    doc = Document(file_path)
    
    root = {
        "id": "root",
        "level": 0,
        "blocks": [],
        "sub_sections": []
    }
    
    section_stack = [root]
    
    counts = {"sec": 0, "p": 0, "t": 0}

    for block in iter_block_items(doc):
        # === A. Handler for Tables ===
        if isinstance(block, Table):
            counts["t"] += 1
            t_id = f"t_{counts['t']}"
            
            # Simple fingerprinting
            table_role = "general_table"
            first_row_text = ""
            if len(block.rows) > 0:
                first_row_cells = [c.text.strip() for c in block.rows[0].cells]
                first_row_text = " ".join(first_row_cells)
                
                if "姓名" in first_row_text and "年级" in first_row_text:
                    table_role = "student_info"
                elif "教学内容" in first_row_text or "教学目标" in first_row_text:
                    table_role = "lesson_meta"
                elif "题目" in first_row_text or "Answer" in first_row_text:
                    table_role = "qa_table"

            table_texts = []
            for row in block.rows:
                for cell in row.cells:
                    cell_text_parts = [cp.text.strip() for cp in cell.paragraphs if cp.text.strip()]
                    if cell_text_parts:
                        table_texts.append(" ".join(cell_text_parts))
            
            block_data = {
                "id": t_id,
                "type": "table",
                "table_role": table_role,
                "rows": len(block.rows),
                "cols": len(block.columns) if block.rows else 0,
                "text_content": table_texts 
            }
            
            section_stack[-1]["blocks"].append(block_data)
            continue

        # === B. Handler for Paragraphs ===
        if isinstance(block, Paragraph):
            style_name = block.style.name if block.style else ""
            text = block.text.strip()
            
            style_info = {
               "name": style_name,
               "alignment": str(block.alignment) if block.alignment else None
            }
            
            is_heading = False
            heading_level = 0
            
            match = re.search(r'(?:Heading|标题)\s*(\d+)', style_name, re.IGNORECASE)
            if match:
                is_heading = True
                heading_level = int(match.group(1))

            if is_heading:
                counts["sec"] += 1
                new_sec_id = f"sec_{counts['sec']}"
                
                sec_role = "general_section"
                if "知识" in text or "讲解" in text:
                    sec_role = "teach"
                elif "训练" in text or "练习" in text:
                    sec_role = "practice"
                elif "回顾" in text:
                    sec_role = "review"
                elif "反思" in text:
                    sec_role = "reflection"
                elif "答案" in text:
                    sec_role = "answer_key"
                
                new_section = {
                    "id": new_sec_id,
                    "title": text,
                    "level": heading_level,
                    "section_role": sec_role,
                    "blocks": [],
                    "sub_sections": []
                }
                
                while section_stack[-1]["level"] >= heading_level:
                    section_stack.pop()
                
                section_stack[-1]["sub_sections"].append(new_section)
                section_stack.append(new_section)
                
            else:
                counts["p"] += 1
                p_id = f"p_{counts['p']}"
                
                block_type = "paragraph"
                payload = {}

                # List detection
                if block._p.pPr is not None and block._p.pPr.numPr is not None:
                    block_type = "list"
                    try:
                        lvl = int(block._p.pPr.numPr.ilvl.val)
                    except (AttributeError, ValueError):
                        lvl = 0
                    payload = {"level": lvl}
                
                elif not text:
                    block_type = "empty_paragraph"
                
                # Question ID detection
                clean_text = text
                if text:
                    q_match = re.match(r'^(\d+)[\.、\s]+(.*)', text)
                    if q_match:
                        payload["question_id"] = int(q_match.group(1))
                        payload["is_question"] = True
                        clean_text = q_match.group(2).strip()

                base_data = {
                    "id": p_id,
                    "type": block_type,
                    "text": clean_text,
                    "raw_text": text,
                    "style": style_info
                }
                base_data.update(payload)
                section_stack[-1]["blocks"].append(base_data)

    # -------------------------------------------------------
    # Post-Processing
    # -------------------------------------------------------
    
    # 1. Recursive filter for empty sections
    def filter_sections(sections):
        valid_sections = []
        for sec in sections:
            sec["sub_sections"] = filter_sections(sec["sub_sections"])
            has_title = bool(sec["title"].strip())
            has_content = len(sec["blocks"]) > 0 or len(sec["sub_sections"]) > 0
            if has_title and has_content:
                valid_sections.append(sec)
        return valid_sections

    root["sub_sections"] = filter_sections(root["sub_sections"])

    # 2. Context Injection
    context_buffer = [] 
    
    def process_blocks(blocks, location_name, parent_role="unknown"):
        nonlocal context_buffer
        processed_list = []
        last_was_empty_slot = False
        
        for block in blocks:
            b_type = block["type"]
            b_text = block.get("raw_text", "").strip()
            b_table_content = block.get("text_content", [])
            
            is_slot = False
            if b_type == "empty_paragraph":
                is_slot = True
            elif b_type == "table":
                is_slot = True
            
            if is_slot:
                if b_type == "empty_paragraph" and last_was_empty_slot:
                    continue 

                block["context"] = list(context_buffer)
                block["location"] = location_name
                block["is_slot"] = True
                
                slot_role = "general_slot"
                if block.get("table_role"):
                    slot_role = block.get("table_role")
                elif parent_role != "unknown":
                    slot_role = f"{parent_role}_content"
                elif "反思" in location_name or "反思" in str(context_buffer):
                     slot_role = "reflection_input"
                
                block["slot_role"] = slot_role
                processed_list.append(block)
                
                if b_type == "empty_paragraph":
                    last_was_empty_slot = True
                else:
                    last_was_empty_slot = False
            else:
                processed_list.append(block)
                last_was_empty_slot = False
            
            if b_text:
                context_buffer.append(b_text)
            if b_table_content:
                for t in b_table_content:
                    context_buffer.append(t)
            if len(context_buffer) > 8:
                context_buffer = context_buffer[-8:]
        
        blocks[:] = processed_list

    if root["blocks"]:
        process_blocks(root["blocks"], "preamble", "preamble")
    
    def traverse_enrich(sections):
        for sec in sections:
            process_blocks(sec["blocks"], sec["title"], sec.get("section_role", "unknown"))
            traverse_enrich(sec["sub_sections"])
            
    traverse_enrich(root["sub_sections"])

    result = {
        "sections": root["sub_sections"]
    }
    if root["blocks"]:
        result["preamble_blocks"] = root["blocks"]
        
    return result

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python parser.py <path_to_docx>")
        sys.exit(1)
        
    file_path = sys.argv[1]
    
    try:
        data = parse_docx(file_path)
        print(json.dumps(data, indent=2, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False))
        sys.exit(1)
