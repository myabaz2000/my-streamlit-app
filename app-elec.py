import streamlit as st
import pandas as pd
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
import os

# =============================================================
# 1. إعدادات الصفحة الرئيسية
# =============================================================
st.set_page_config(
    page_title="منظومة إدارة مكاتب التصويت - قيادة أملن",
    page_icon="🗳️",
    layout="wide"
)

# =============================================================
# 2. مسارات الملفات (مسارات نسبية للعمل على Streamlit Cloud)
# =============================================================
EXCEL_FILE = 'Bureaux_de_Votes_Ammelne_2026_Modf_2026.xlsx'
LOGISTICS_FILE = 'logistics_data.xlsx'
TRANSPORT_FILE = 'transport_data.xlsx'

# =============================================================
# 3. دالة تحميل البيانات بأمان
# =============================================================
@st.cache_data
def load_data():
    offices_dict = {}
    
    # تحميل الملف الرئيسي (مكاتب التصويت)
    if os.path.exists(EXCEL_FILE):
        try:
            excel_data = pd.ExcelFile(EXCEL_FILE)
            for sheet in excel_data.sheet_names:
                df = pd.read_excel(excel_data, sheet_name=sheet)
                df.columns = [str(col).strip().lower() for col in df.columns]
                
                # إضافة عمود الجماعة إذا لم يكن موجوداً
                if 'commune' not in df.columns:
                    df['commune'] = sheet
                    
                offices_dict[sheet] = df
        except Exception as e:
            st.error(f"خطأ أثناء قراءة ملف المكاتب: {e}")
            
    # دمج بيانات الجماعات في جدول واحد
    if offices_dict:
        df_offices = pd.concat(offices_dict.values(), ignore_index=True)
    else:
        df_offices = pd.DataFrame()

    # تحميل ملف اللوجستيك
    if os.path.exists(LOGISTICS_FILE):
        try:
            df_logistics = pd.read_excel(LOGISTICS_FILE)
            df_logistics.columns = [str(col).strip().lower() for col in df_logistics.columns]
        except Exception:
            df_logistics = pd.DataFrame()
    else:
        df_logistics = pd.DataFrame()

    # تحميل ملف النقل والوسائل
    if os.path.exists(TRANSPORT_FILE):
        try:
            df_transport = pd.read_excel(TRANSPORT_FILE)
            df_transport.columns = [str(col).strip().lower() for col in df_transport.columns]
        except Exception:
            df_transport = pd.DataFrame()
    else:
        df_transport = pd.DataFrame()

    return df_offices, df_logistics, df_transport, list(offices_dict.keys())

# تحميل البيانات
df_offices, df_logistics, df_transport, list_communes = load_data()

# =============================================================
# 4. الواجهة والترووس (Header)
# =============================================================
st.title("🗳️ منظومة إدارة مكاتب التصويت - قيادة أملن")
st.subheader("الجماعات: أملن - تارسواط - تاسريرت (رؤساء المكاتب ونوابهم والأعضاء ونوابهم)")

# الشريط الجانبي للتصفية
st.sidebar.header("🔍 خيارات التصفية")

if list_communes:
    selected_commune = st.sidebar.selectbox("اختر الجماعة:", ["الكل"] + list_communes)
else:
    selected_commune = "الكل"

# تصفية البيانات حسب الجماعة
if selected_commune != "الكل" and not df_offices.empty and "commune" in df_offices.columns:
    filtered_offices = df_offices[df_offices["commune"] == selected_commune]
else:
    filtered_offices = df_offices.copy()

# =============================================================
# 5. إدارة التبويبات (Tabs)
# =============================================================
tabs = [
    "👥 رؤساء المكاتب ونوابهم والأعضاء",
    "📇 البطاقة التقنية",
    "🏢 التجهيزات واللوجستيك",
    "🚗 التنقل والوسائل",
    "🗺️ الخريطة والمواقع"
]

if "current_tab" not in st.session_state:
    st.session_state.current_tab = tabs[0]

selected_tab = st.radio("", tabs, horizontal=True)
st.session_state.current_tab = selected_tab

st.markdown("---")

# -------------------------------------------------------------
# TAB 1: رؤساء المكاتب ونوابهم والأعضاء
# -------------------------------------------------------------
if st.session_state.current_tab == "👥 رؤساء المكاتب ونوابهم والأعضاء":
    st.header(f"📋 لائحة المكاتب والأعضاء ({selected_commune})")
    
    if not filtered_offices.empty:
        st.dataframe(filtered_offices, use_container_width=True)
    else:
        st.info("لا توجد بيانات متوفرة للعرض.")

# -------------------------------------------------------------
# TAB 2: البطاقة التقنية
# -------------------------------------------------------------
elif st.session_state.current_tab == "📇 البطاقة التقنية":
    st.header("📇 البحث عن البطاقة التقنية لمكتب تصويت")
    
    col1, col2 = st.columns(2)
    with col1:
        card_commune = st.selectbox("الجماعة:", list_communes if list_communes else ["لا يوجد"])
    with col2:
        card_office_num = st.number_input("رقم المكتب:", min_value=1, value=1, step=1)
        
    # الحماية من أخطاء KeyError وتصفية الصف بأمان
    if not df_offices.empty and "commune" in df_offices.columns and "num" in df_offices.columns:
        info_row = df_offices[(df_offices["commune"] == card_commune) & (df_offices["num"].astype(str) == str(card_office_num))]
    else:
        info_row = pd.DataFrame()
        
    if not info_row.empty:
        st.success(f"تم العثور على بيانات المكتب رقم {card_office_num} - جماعة {card_commune}")
        st.json(info_row.iloc[0].to_dict())
    else:
        st.warning("لا توجد بيانات متوفرة لهذا المكتب.")

# -------------------------------------------------------------
# TAB 3: التجهيزات واللوجستيك
# -------------------------------------------------------------
elif st.session_state.current_tab == "🏢 التجهيزات واللوجستيك":
    st.header(f"🏢 بيانات اللوجستيك والتجهيزات ({selected_commune})")
    
    if not df_logistics.empty:
        st.dataframe(df_logistics, use_container_width=True)
    else:
        st.info("لا توجد بيانات لوجستية متوفرة.")

# -------------------------------------------------------------
# TAB 4: التنقل والوسائل
# -------------------------------------------------------------
elif st.session_state.current_tab == "🚗 التنقل والوسائل":
    st.header(f"🚗 وسائل النقل والتنقلات ({selected_commune})")
    
    if not df_transport.empty:
        st.dataframe(df_transport, use_container_width=True)
    else:
        st.info("لا توجد بيانات متوفرة لوسائل النقل.")

# -------------------------------------------------------------
# TAB 5: الخريطة والمواقع
# -------------------------------------------------------------
elif st.session_state.current_tab == "🗺️ الخريطة والمواقع":
    st.header(f"🗺️ خريطة مواقع مكاتب التصويت ({selected_commune})")
    
    if not filtered_offices.empty and "lat" in filtered_offices.columns and "lon" in filtered_offices.columns:
        map_df = filtered_offices.copy()
        
        # تنظيف وتحويل الإحداثيات إلى أرقام
        map_df['lat'] = map_df['lat'].astype(str).str.replace(',', '.')
        map_df['lon'] = map_df['lon'].astype(str).str.replace(',', '.')
        
        map_df['lat'] = pd.to_numeric(map_df['lat'], errors='coerce')
        map_df['lon'] = pd.to_numeric(map_df['lon'], errors='coerce')
        
        # تصفية النقاط الصالحة
        valid_offices = map_df.dropna(subset=["lat", "lon"])
        
        if not valid_offices.empty:
            avg_lat = valid_offices["lat"].mean()
            avg_lon = valid_offices["lon"].mean()
            
            # إنشاء الخريطة
            m = folium.Map(
                location=[avg_lat, avg_lon], 
                zoom_start=11
            )
            
            folium.TileLayer('OpenStreetMap', name='الخريطة العادية').add_to(m)
            
            marker_cluster = MarkerCluster().add_to(m)
            for _, row in valid_offices.iterrows():
                lat, lon = float(row["lat"]), float(row["lon"])
                
                popup_html = f"""
                <div dir="rtl" style="font-family: Arial; text-align: right;">
                    <b>جماعة {row.get('commune', '')}</b><br>
                    مكتب رقم: {row.get('num', '')}<br>
                    {row.get('center_name', row.get('name', ''))}
                </div>
                """
                
                folium.Marker(
                    [lat, lon],
                    popup=folium.Popup(popup_html, max_width=250),
                    tooltip=f"مكتب {row.get('num', '')} - جماعة {row.get('commune', '')}"
                ).add_to(marker_cluster)
            
            st_folium(m, width="100%", height=550, key="main_map_display")
            
        else:
            st.error("❌ لم يتم العثور على إحداثيات (lat/lon) صالحة في الملف.")
    else:
        st.warning("⚠️ لا توجد أعمدة إحداثيات متوفرة (lat/lon) في ملف Excel الرئيسي.")
