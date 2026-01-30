import json
import sys
import os

def generate_content(structure_path, topic, output_path):
    print(f"Loading structure from: {structure_path}")
    print(f"Topic: {topic}")
    
    with open(structure_path, 'r') as f:
        data = json.load(f)
        
    content_map = {}
    
    # Recursive traversal to find slots
    def traverse(sections):
        for sec in sections:
            # Check blocks in this section
            for block in sec.get("blocks", []):
                process_block(block)
            
            # Recurse
            traverse(sec.get("sub_sections", []))
            
    def process_block(block):
        if not block.get("is_slot"):
            return
            
        b_id = block["id"]
        role = block.get("slot_role", "general")
        loc = block.get("location", "unknown")
        
        # --- Generation Logic (Mock/Rule-Based) ---
        # In a real scenario, this is where you'd call an LLM API 
        # passing the 'context' and 'role' from the block.
        
        generated_text = ""
        
        if block["type"] == "table":
            # For table, we need a list of strings (one per row, usually for the content column)
            row_count = block.get("rows", 0)
            
            if role == "student_info":
                 generated_text = ["XXX" for _ in range(row_count)]
            elif role == "teach_content":
                generated_text = [
                    f"【Concept {i+1}】 Detailed explanation about {topic}..." 
                    for i in range(row_count)
                ]
            else:
                generated_text = [
                    f"[{role} Table Content] Row {i+1} for {topic}" 
                    for i in range(row_count)
                ]
        else:
            # Paragraph
            if role == "teach_content":
                generated_text = f"Here is a detailed explanation of {topic} in the context of {loc}."
            elif role == "practice_content":
                generated_text = f"1. Question about {topic}?\n2. Another question about {topic}?"
            elif role == "reflection_input":
                generated_text = "(Student reflection space)"
            else:
                 generated_text = f"[{role}] Generated content for {topic}."

        content_map[b_id] = generated_text

    # Start traversal
    # 1. Preamble
    if "preamble_blocks" in data:
        for b in data["preamble_blocks"]:
            process_block(b)
            
    # 2. Sections
    traverse(data.get("sections", []))
    
    print(f"Generated {len(content_map)} content items.")
    
    # Save
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(content_map, f, indent=2, ensure_ascii=False)
    print(f"Saved content map to {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python generator.py <structure_json> <topic> <output_json>")
        sys.exit(1)
        
    struct_file = sys.argv[1]
    topic_str = sys.argv[2]
    out_file = sys.argv[3]
    
    generate_content(struct_file, topic_str, out_file)
