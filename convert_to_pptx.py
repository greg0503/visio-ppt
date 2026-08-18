import win32com.client
import os
import time
import zipfile
import re

def safe_copy_paste(visio, vpage, stencil, m_name, slide, x, y, card_w, card_h, max_retries=5):
    m = stencil.Masters.Item(m_name)
    v_shape = vpage.Drop(m, 4.0, 4.0)
    
    ungrouped = None
    for attempt in range(max_retries):
        try:
            time.sleep(0.08 * (attempt + 1))
            visio.ActiveWindow.Select(v_shape, 2)
            visio.ActiveWindow.Selection.Copy()
            time.sleep(0.08 * (attempt + 1))
            
            p_shape = slide.Shapes.PasteSpecial(DataType=2) # 2 = ppPasteEnhancedMetafile
            ungrouped = p_shape.Ungroup()
            break
        except Exception as e:
            if attempt == max_retries - 1:
                print(f"Failed to copy-paste master '{m_name}' after {max_retries} attempts: {e}")
                try:
                    v_shape.Delete()
                except:
                    pass
                return False
            time.sleep(0.15)

    if ungrouped is not None:
        try:
            max_w = card_w - 24
            max_h = card_h - 40
            if ungrouped.Width > max_w or ungrouped.Height > max_h:
                scale = min(max_w / ungrouped.Width, max_h / ungrouped.Height)
                ungrouped.LockAspectRatio = -1 # True
                ungrouped.Width = ungrouped.Width * scale

            ungrouped.Left = x + (card_w - ungrouped.Width) / 2
            ungrouped.Top = y + 8 + (max_h - ungrouped.Height) / 2
        except Exception as e:
            print(f"Layout warning for '{m_name}': {e}")

    try:
        v_shape.Delete()
    except:
        pass
        
    return True

def sanitize_metadata(pptx_path):
    print("Sanitizing presentation metadata and personal traces...")
    temp_zip_path = pptx_path + ".tmp.zip"
    
    with zipfile.ZipFile(pptx_path, 'r') as zin:
        with zipfile.ZipFile(temp_zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                
                # Sanitize docProps/core.xml
                if item.filename == "docProps/core.xml":
                    xml_str = data.decode('utf-8', errors='ignore')
                    xml_str = re.sub(r'<dc:creator>.*?</dc:creator>', '<dc:creator>TSirc Group</dc:creator>', xml_str)
                    xml_str = re.sub(r'<cp:lastModifiedBy>.*?</cp:lastModifiedBy>', '<cp:lastModifiedBy>TSirc Group</cp:lastModifiedBy>', xml_str)
                    xml_str = re.sub(r'<dc:description>.*?</dc:description>', '<dc:description>Circuit Stencil for PowerPoint</dc:description>', xml_str)
                    data = xml_str.encode('utf-8')
                    
                # Sanitize docProps/app.xml
                elif item.filename == "docProps/app.xml":
                    xml_str = data.decode('utf-8', errors='ignore')
                    xml_str = re.sub(r'<Company>.*?</Company>', '<Company></Company>', xml_str)
                    xml_str = re.sub(r'<Manager>.*?</Manager>', '<Manager></Manager>', xml_str)
                    data = xml_str.encode('utf-8')
                    
                zout.writestr(item, data)
                
    os.remove(pptx_path)
    os.rename(temp_zip_path, pptx_path)
    print("Metadata successfully sanitized!")

def build_stencil_presentation():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    vssx_path = os.path.join(base_dir, 'Kwantae_Circuit.vssx')
    output_pptx = os.path.join(base_dir, 'Kwantae_Circuit_Stencil.pptx')
    
    if os.path.exists(output_pptx):
        try:
            os.remove(output_pptx)
        except Exception as e:
            print("Notice: Existing output file removal:", e)

    categories = [
        {
            "category_title": "Transistors (MOSFETs & BJTs)",
            "subtitle": "CMOS Transistors (NMOS, PMOS, Body-Tie, Diode-connected) & Bipolar Junction Transistors",
            "masters": [
                "NMOS", "PMOS",
                "LVT NMOS", "LVT PMOS",
                "HVT NMOS", "HVT PMOS",
                "NMOS BodyTie", "PMOS BodyTie",
                "NMOS Diode", "PMOS Diode",
                "NPN", "PNP"
            ],
            "cols": 4,
            "card_w": 205,
            "card_h": 125,
            "start_x": 40,
            "start_y": 105
        },
        {
            "category_title": "Passive Components & Power Rails",
            "subtitle": "Resistors, Capacitors, Inductors, VDD/GND Supplies, Nodes, Switches & Basic Elements",
            "masters": [
                "Resistor", "Inductor", "Capacitor", "Diode",
                "VDD", "GND", "Voltage Node", "Small Node",
                "Switch", "Chopper", "Current Source", "PAD",
                "Crystal", "Wire"
            ],
            "cols": 5,
            "card_w": 165,
            "card_h": 120,
            "start_x": 40,
            "start_y": 105
        },
        {
            "category_title": "Amplifiers & Analog Functional Blocks",
            "subtitle": "Operational Amplifiers, Fully-Differential Op-Amps, Comparators & Analog Blocks",
            "masters": [
                "Amp", "FD OpAmp", "Comparator",
                "TC", "TI", "Electrode"
            ],
            "cols": 3,
            "card_w": 275,
            "card_h": 160,
            "start_x": 40,
            "start_y": 115
        },
        {
            "category_title": "Digital Logic Gates & Flip-Flops",
            "subtitle": "Inverters, Basic Logic Gates (AND/NAND/OR/NOR/XOR) & D Flip-Flops",
            "masters": [
                "INV", "AND", "NAND", "OR",
                "NOR", "XOR", "D FF Small", "D FF Large"
            ],
            "cols": 4,
            "card_w": 205,
            "card_h": 160,
            "start_x": 40,
            "start_y": 115
        },
        {
            "category_title": "Signal Sources & Waveforms",
            "subtitle": "Clock, Sine Waves, Non-overlapping Clocks, Pulse & Noise Sources",
            "masters": [
                "Clock", "Pure-Sine", "Pseudo-Sine", "Light",
                "Non-Overlap", "ECG", "APW", "Noise"
            ],
            "cols": 4,
            "card_w": 205,
            "card_h": 160,
            "start_x": 40,
            "start_y": 115
        },
        {
            "category_title": "Biomedical & System Level Symbols",
            "subtitle": "Biomedical Sensor Interfaces, Anatomical & System Graphics",
            "masters": [
                "Heart", "Vessel", "Chest",
                "Upper Body", "Human", "Hand"
            ],
            "cols": 3,
            "card_w": 275,
            "card_h": 160,
            "start_x": 40,
            "start_y": 115
        }
    ]

    print("Initializing Visio & PowerPoint...")
    visio = win32com.client.Dispatch('Visio.Application')
    visio.Visible = False
    stencil = visio.Documents.OpenEx(vssx_path, 2)
    vdoc = visio.Documents.Add('')
    vpage = vdoc.Pages.Item(1)

    ppt = win32com.client.Dispatch('PowerPoint.Application')
    ppt.Visible = True
    pres = ppt.Presentations.Add(True)
    
    # 16:9 Widescreen (960 x 540 pt)
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540

    def rgb(r, g, b):
        return r + (g * 256) + (b * 65536)

    NAVY_PRIMARY = rgb(26, 54, 93)     # #1A365D
    GRAY_TEXT    = rgb(100, 116, 139) # #64748B
    CARD_BG      = rgb(248, 250, 252) # #F8FAFC
    CARD_BORDER  = rgb(226, 232, 240) # #E2E8F0
    ACCENT_BLUE  = rgb(37, 99, 235)   # #2563EB
    LABEL_COLOR  = rgb(30, 41, 59)    # #1E293B

    # ==========================================
    # Slide 1: Cover & User Guide Slide (100% English)
    # ==========================================
    cover_slide = pres.Slides.Add(1, 12)
    
    banner = cover_slide.Shapes.AddShape(1, 0, 0, 960, 8)
    banner.Fill.Solid()
    banner.Fill.ForeColor.RGB = ACCENT_BLUE
    banner.Line.Visible = False

    t_box = cover_slide.Shapes.AddTextbox(1, 60, 50, 840, 60)
    tf = t_box.TextFrame
    tf.TextRange.Text = "Circuit Stencil for PowerPoint"
    tf.TextRange.Font.Name = "Segoe UI"
    tf.TextRange.Font.Size = 28
    tf.TextRange.Font.Bold = True
    tf.TextRange.Font.Color.RGB = NAVY_PRIMARY

    st_box = cover_slide.Shapes.AddTextbox(1, 60, 110, 840, 30)
    stf = st_box.TextFrame
    stf.TextRange.Text = "High-Quality Vector Circuit Components (Analog IC, Razavi & TSirc Style)"
    stf.TextRange.Font.Name = "Segoe UI"
    stf.TextRange.Font.Size = 14
    stf.TextRange.Font.Color.RGB = ACCENT_BLUE

    guide_card = cover_slide.Shapes.AddShape(1, 60, 160, 840, 330)
    guide_card.Fill.Solid()
    guide_card.Fill.ForeColor.RGB = CARD_BG
    guide_card.Line.ForeColor.RGB = CARD_BORDER
    guide_card.Line.Weight = 1.0

    guide_text = cover_slide.Shapes.AddTextbox(1, 80, 180, 800, 290)
    gtf = guide_text.TextFrame
    gtf.WordWrap = True
    guide_content = (
        "💡 Quick User Guide:\n\n"
        "1. 100% Native Vector Shapes (Fully Editable):\n"
        "   - All schematic symbols are converted into native PowerPoint Drawing Shapes/Groups.\n"
        "   - Infinite scalability without loss of resolution. Line weights, colors, and fills are fully customizable.\n\n"
        "2. Copy & Paste:\n"
        "   - Simply select any component, copy (Ctrl + C), and paste (Ctrl + V) into your own presentation slides.\n\n"
        "3. Rotation & Orientation:\n"
        "   - For mirrored transistors or flipped rails, use [Shape Format] -> [Rotate] -> [Flip Horizontal / Vertical].\n\n"
        "4. Categorization (Slides 2 - 7):\n"
        "   - Slide 2: Transistors (NMOS, PMOS, Body-Tie, Diode-connected, BJTs)\n"
        "   - Slide 3: Passive Components, Power Supplies & Basic Nodes\n"
        "   - Slide 4: Operational Amplifiers, Fully-Differential Op-Amps & Comparators\n"
        "   - Slide 5: Digital Logic Gates & D Flip-Flops\n"
        "   - Slide 6: Clock Sources, Sine Waveforms & Noise\n"
        "   - Slide 7: Biomedical & System Level Icons"
    )
    gtf.TextRange.Text = guide_content
    gtf.TextRange.Font.Name = "Segoe UI"
    gtf.TextRange.Font.Size = 12
    gtf.TextRange.Font.Color.RGB = LABEL_COLOR

    # ==========================================
    # Slides 2-7: Category Stencil Slides
    # ==========================================
    for cat_idx, cat in enumerate(categories):
        slide_num = cat_idx + 2
        slide = pres.Slides.Add(slide_num, 12)

        top_bar = slide.Shapes.AddShape(1, 0, 0, 960, 4)
        top_bar.Fill.Solid()
        top_bar.Fill.ForeColor.RGB = ACCENT_BLUE
        top_bar.Line.Visible = False

        h_box = slide.Shapes.AddTextbox(1, 40, 18, 880, 35)
        htf = h_box.TextFrame
        htf.TextRange.Text = f"{cat['category_title']}"
        htf.TextRange.Font.Name = "Segoe UI"
        htf.TextRange.Font.Size = 18
        htf.TextRange.Font.Bold = True
        htf.TextRange.Font.Color.RGB = NAVY_PRIMARY

        sub_box = slide.Shapes.AddTextbox(1, 40, 52, 880, 25)
        sbtf = sub_box.TextFrame
        sbtf.TextRange.Text = cat['subtitle']
        sbtf.TextRange.Font.Name = "Segoe UI"
        sbtf.TextRange.Font.Size = 10.5
        sbtf.TextRange.Font.Color.RGB = GRAY_TEXT

        cols = cat['cols']
        card_w = cat['card_w']
        card_h = cat['card_h']
        start_x = cat['start_x']
        start_y = cat['start_y']
        gap_x = 12
        gap_y = 12

        print(f"Building slide {slide_num}: {cat['category_title']} ({len(cat['masters'])} shapes)...")

        for i, m_name in enumerate(cat['masters']):
            row = i // cols
            col = i % cols
            x = start_x + col * (card_w + gap_x)
            y = start_y + row * (card_h + gap_y)

            card = slide.Shapes.AddShape(1, x, y, card_w, card_h)
            card.Fill.Solid()
            card.Fill.ForeColor.RGB = CARD_BG
            card.Line.ForeColor.RGB = CARD_BORDER
            card.Line.Weight = 0.75

            lbl = slide.Shapes.AddTextbox(1, x, y + card_h - 26, card_w, 22)
            ltf = lbl.TextFrame
            ltf.TextRange.Text = m_name
            ltf.TextRange.Font.Name = "Segoe UI Semibold"
            ltf.TextRange.Font.Size = 10
            ltf.TextRange.Font.Color.RGB = LABEL_COLOR
            ltf.TextRange.ParagraphFormat.Alignment = 2

            safe_copy_paste(visio, vpage, stencil, m_name, slide, x, y, card_w, card_h)

    # Save Presentation
    pres.SaveCopyAs(output_pptx)
    print(f"Presentation successfully saved to: {output_pptx}")

    # Safe Cleanup
    try:
        pres.Saved = True
        pres.Close()
    except Exception as e:
        print("PPT Close note:", e)
    try:
        ppt.Quit()
    except Exception as e:
        print("PPT Quit note:", e)
    try:
        vdoc.Saved = True
        vdoc.Close()
    except Exception as e:
        print("Visio Close note:", e)
    try:
        stencil.Close()
    except Exception as e:
        print("Stencil Close note:", e)
    try:
        visio.Quit()
    except Exception as e:
        print("Visio Quit note:", e)
    
    # Sanitize metadata
    sanitize_metadata(output_pptx)
    
    return output_pptx

if __name__ == '__main__':
    build_stencil_presentation()
