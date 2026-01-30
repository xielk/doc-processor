import json
import sys
import os
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

# Re-use iteration logic to ensure ID alignment
def iter_block_items(parent):
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

def build_doc(template_path, content_json_path, output_path):
    print(f"Loading template: {template_path}")
    
    # 检查模板文件是否存在
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"\n❌ 模板文件不存在: {template_path}\n"
            f"\n可能原因：\n"
            f"1. 文件路径错误\n"
            f"2. 文件在当前session结束后被清理（/tmp/目录下的文件会被定期清理）\n"
            f"3. 文件被移动或删除\n"
            f"\n解决方案：\n"
            f"1. 请提供正确的模板文件路径（绝对路径）\n"
            f"2. 将文件复制到非/tmp/目录，如: /Users/xielk/webdata/english/lesson/templates/\n"
            f"3. 重新上传模板文件\n"
        )
    
    doc = Document(template_path)
    
    print(f"Loading content: {content_json_path}")
    
    # 检查内容文件是否存在
    if not os.path.exists(content_json_path):
        raise FileNotFoundError(
            f"\n❌ 内容文件不存在: {content_json_path}\n"
            f"\n请先运行内容生成流程，创建content.json文件\n"
        )
    
    with open(content_json_path, 'r') as f:
        content_map = json.load(f)
        
    # Counters to match parser.py IDs
    counts = {"p": 0, "t": 0, "sec": 0} # sec count is trickier as it depends on heading logic.
    # Note: parser.py tracks sections via Heading Styles. We must replicate that logic to sync IDs.
    # However, parser.py assigns IDs to blocks (`p_X`, `t_X`) independently of section hierarchy?
    # Wait, parser.py does: `counts["p"] += 1` inside the loop.
    # `counts["sec"]` is for sections.
    # The Block IDs are `p_{count}` or `t_{count}`.
    # So we simply need to iterate and increment counters same way.
    
    import re
    
    for block in iter_block_items(doc):
        # Handle Table
        if isinstance(block, Table):
            counts["t"] += 1
            t_id = f"t_{counts['t']}"
            
            if t_id in content_map:
                # Fill table
                # The content_map usually provides "text_content" list or specific cell updates?
                # For simplicity, if content_map[t_id] is a string, we might put it in first cell?
                # Or structure: {"row_x_col_y": "text"}?
                # Let's assume content_map[t_id] is a list of strings to fill the "Content Column" (Col 1)?
                
                data = content_map[t_id]
                if isinstance(data, list):
                    # Smart fill: Try to fill 2nd column (index 1) of each row?
                    # Or just fill sequentially?
                    # Let's assume data is a list matching rows.
                    for i, text in enumerate(data):
                        if i < len(block.rows):
                            # Try to write to Col 1, else Col 0
                            target_cell_idx = 1 if len(block.rows[i].cells) > 1 else 0
                            block.rows[i].cells[target_cell_idx].text = str(text)
            continue
            
        # Handle Paragraph
        if isinstance(block, Paragraph):
            style_name = block.style.name if block.style else ""
            
            # Check for Heading (Section detection)
            # This logic must match parser.py purely to skip "p" counter for headings?
            # Let's check parser.py:
            # if is_heading: counts["sec"] +=1 ... (Does NOT increment p)
            # else: counts["p"] += 1 ...
            
            is_heading = False
            match = re.search(r'(?:Heading|标题)\s*(\d+)', style_name, re.IGNORECASE)
            if match:
                is_heading = True
            
            if is_heading:
                counts["sec"] += 1
                # We usually don't replace headers, but maybe?
                # sec_id = f"sec_{counts['sec']}"
                # if sec_id in content_map: block.text = content_map[sec_id]
            else:
                counts["p"] += 1
                p_id = f"p_{counts['p']}"
                
                if p_id in content_map:
                    # Update text
                    block.text = str(content_map[p_id])

    doc.save(output_path)
    print(f"Generated doc saved to {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python builder.py <template_docx> <content.json> <output_docx>")
        print("\n示例:")
        print("  python builder.py /path/to/template.docx /path/to/content.json /path/to/output.docx")
        sys.exit(1)
        
    tpl = sys.argv[1]
    content = sys.argv[2]
    out = sys.argv[3]
    
    try:
        build_doc(tpl, content, out)
        print("\n✅ 文档生成成功!")
    except FileNotFoundError as e:
        print(f"\n{e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 生成文档时出错: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
