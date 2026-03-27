import streamlit as st
import cv2
import easyocr
import numpy as np
import io
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

# ---------------------------------------------------------
# 1. 設定網頁標題與外觀
# ---------------------------------------------------------
st.set_page_config(page_title="圖片轉可編輯 PPT 神器", page_icon="🪄")
st.title("🪄 圖片轉可編輯 PPT 神器")
st.write("上傳一張圖片，AI 會自動將裡面的圖案轉成「積木色塊」，並把文字轉成「PPT 文字方塊」！")

# ---------------------------------------------------------
# 2. 載入 AI 模型 (加入 Cache 讓它只載入一次，不用每次等)
# ---------------------------------------------------------
@st.cache_resource
def load_ocr_model():
    return easyocr.Reader(['ch_tra', 'en'], gpu=False)

reader = load_ocr_model()

# ---------------------------------------------------------
# 3. 網頁控制面板 (側邊欄)
# ---------------------------------------------------------
with st.sidebar:
    st.header("⚙️ 轉換設定")
    resolution = st.selectbox(
        "選擇圖案精細度 (解析度)",
        ('Low (快, 150區塊)', 'Medium (適中, 250區塊)', 'High (慢, 350區塊)')
    )
    # 抓取對應的數值
    if 'Low' in resolution: res_val = 150
    elif 'Medium' in resolution: res_val = 250
    else: res_val = 350

# ---------------------------------------------------------
# 4. 圖片上傳區塊
# ---------------------------------------------------------
uploaded_file = st.file_uploader("📂 請上傳圖片 (支援 PNG, JPG, JPEG)", type=['png', 'jpg', 'jpeg'])

if uploaded_file is not None:
    # 在網頁上顯示使用者上傳的圖片
    st.image(uploaded_file, caption="上傳的圖片預覽", use_container_width=True)
    
    # 建立一個轉換按鈕
    if st.button("🚀 開始轉換為 PPT"):
        
        # 建立進度條與狀態文字
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.text("🧠 正在讀取圖片與辨識文字...")
        
        # 將上傳的檔案轉換為 OpenCV 看得懂的格式
        file_bytes = np.asarray(bytearray(uploaded_file.read()), dtype=np.uint8)
        img = cv2.imdecode(file_bytes, 1)
        img_h, img_w = img.shape[:2]

        # 準備 PPT
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # === 【AI 文字辨識與虛擬橡皮擦】 ===
        text_results = reader.readtext(img)
        text_data_to_draw = []

        if text_results:
            for (bbox, text, prob) in text_results:
                text_data_to_draw.append((bbox, text))
                pts = np.array(bbox, np.int32)
                cv2.fillPoly(img, [pts], (255, 255, 255)) # 把文字塗白

        # === 【繪製圖形 (掃描線積木填色)】 ===
        status_text.text("🎨 正在繪製積木色塊...")
        ratio = res_val / img_w
        block_h = int(img_h * ratio)
        small_img = cv2.resize(img, (res_val, block_h), interpolation=cv2.INTER_AREA)

        max_ppt_w, max_ppt_h = 720, 540
        ppt_scale = min(max_ppt_w / res_val, max_ppt_h / block_h) * 0.8
        offset_x = (max_ppt_w - (res_val * ppt_scale)) / 2
        offset_y = (max_ppt_h - (block_h * ppt_scale)) / 2

        for y in range(block_h):
            # 更新網頁上的進度條
            progress_bar.progress((y + 1) / block_h)
            
            x = 0
            while x < res_val:
                b, g, r = small_img[y, x]
                if r > 240 and g > 240 and b > 240:
                    x += 1
                    continue
                    
                start_x = x
                while x < res_val:
                    nb, ng, nr = small_img[y, x]
                    if abs(int(nr)-int(r)) <= 3 and abs(int(ng)-int(g)) <= 3 and abs(int(nb)-int(b)) <= 3:
                        x += 1
                    else:
                        break
                width = x - start_x
                
                ppt_x = start_x * ppt_scale + offset_x
                ppt_y = y * ppt_scale + offset_y
                ppt_width = width * ppt_scale
                ppt_height = ppt_scale
                
                rect = slide.shapes.add_shape(1, Pt(ppt_x), Pt(ppt_y), Pt(ppt_width + 0.3), Pt(ppt_height + 0.3))
                rect.fill.solid()
                rect.fill.fore_color.rgb = RGBColor(int(r), int(g), int(b))
                rect.line.fill.background() 

        # === 【在圖形上疊加文字方塊】 ===
        status_text.text("📝 正在排版文字方塊...")
        for bbox, text in text_data_to_draw:
            orig_x, orig_y = bbox[0][0], bbox[0][1]
            orig_w, orig_h = bbox[1][0] - bbox[0][0], bbox[2][1] - bbox[1][1]
            
            ppt_text_x = (orig_x / img_w) * (res_val * ppt_scale) + offset_x
            ppt_text_y = (orig_y / img_h) * (block_h * ppt_scale) + offset_y
            ppt_text_w = max((orig_w / img_w) * (res_val * ppt_scale), 10)
            ppt_text_h = max((orig_h / img_h) * (block_h * ppt_scale), 10)
            
            txBox = slide.shapes.add_textbox(Pt(ppt_text_x), Pt(ppt_text_y), Pt(ppt_text_w), Pt(ppt_text_h))
            tf = txBox.text_frame
            tf.text = text
            
            font_size = max(int(ppt_text_h * 0.7), 10)
            tf.paragraphs[0].font.size = Pt(font_size)
            tf.paragraphs[0].font.color.rgb = RGBColor(255, 0, 0)

        # === 【將 PPT 存入記憶體並提供下載】 ===
        status_text.text("✅ 轉換完成！請點擊下方按鈕下載 PPT。")
        
        # 這裡不存成實體檔案，而是存在「記憶體(BytesIO)」中，這樣網頁才能提供下載
        ppt_stream = io.BytesIO()
        prs.save(ppt_stream)
        ppt_stream.seek(0)
        
        # 顯示綠色的成功訊息與下載按鈕
        st.success("轉換成功！")
        st.download_button(
            label="📥 下載可編輯的 PPT 檔案",
            data=ppt_stream,
            file_name="converted_editable.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )