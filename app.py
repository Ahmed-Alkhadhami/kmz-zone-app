# ======================================================
# الاستيرادات
# ======================================================
import io
import zipfile
import json
import random
import streamlit as st
import pandas as pd

from shapely.geometry import Polygon, Point
from shapely.strtree import STRtree
from bs4 import BeautifulSoup
from xml.dom.minidom import Document
import xml.etree.ElementTree as ET


# ======================================================
# إعداد الصفحة
# ======================================================
st.set_page_config(
    page_title="🗺️ أداة استخراج وفحص الزونات",
    layout="wide"
)

st.title("🗺️ أداة استخراج وفحص الزونات (KMZ / KML / Excel)")


# ======================================================
# دوال مساعدة
# ======================================================

def random_kml_color(fill_alpha="55", line_alpha="FF"):
    """توليد لون عشوائي لـ KML"""
    r, g, b = [random.randint(0, 255) for _ in range(3)]
    return (
        f"{fill_alpha}{b:02x}{g:02x}{r:02x}",
        f"{line_alpha}{b:02x}{g:02x}{r:02x}"
    )

def calculate_area_in_sqm(polygon):
    """حساب مساحة البوليقون بالمتر المربع باستخدام تقريب"""
    # تحويل من درجات إلى متر (تقريبي)
    # عند خط الاستواء: 1 درجة ≈ 111 كم
    # نستخدم نقطة المنتصف لحساب أدق
    centroid = polygon.centroid
    lat = centroid.y
    
    # عامل التحويل للطول (longitude) يعتمد على خط العرض
    # 1 درجة طول = 111320 * cos(lat) متر
    # 1 درجة عرض = 111320 متر
    import math
    lon_factor = 111320 * math.cos(math.radians(lat))
    lat_factor = 111320
    
    # حساب المساحة بالدرجات المربعة ثم تحويلها
    area_deg = polygon.area
    area_sqm = area_deg * lon_factor * lat_factor
    
    return area_sqm


def calculate_distance_in_meters(point1, point2):
    """حساب المسافة بين نقطتين بالمتر"""
    import math
    
    # استخراج الإحداثيات
    if isinstance(point1, Point):
        lon1, lat1 = point1.x, point1.y
    else:
        lon1, lat1 = point1
        
    if isinstance(point2, Point):
        lon2, lat2 = point2.x, point2.y
    else:
        lon2, lat2 = point2
    
    # حساب المسافة باستخدام صيغة Haversine
    R = 6371000  # نصف قطر الأرض بالمتر
    
    lat1_rad = math.radians(lat1)
    lat2_rad = math.radians(lat2)
    delta_lat = math.radians(lat2 - lat1)
    delta_lon = math.radians(lon2 - lon1)
    
    a = math.sin(delta_lat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(delta_lon/2)**2
    c = 2 * math.asin(math.sqrt(a))
    
    distance = R * c
    return distance

def parse_kmz_or_kml(uploaded_file):
    """قراءة KMZ أو KML واستخراج البوليقونز"""
    if uploaded_file.name.lower().endswith(".kmz"):
        with zipfile.ZipFile(uploaded_file, "r") as kmz:
            kml_name = [f for f in kmz.namelist() if f.endswith(".kml")][0]
            kml_data = kmz.read(kml_name)
    else:
        kml_data = uploaded_file.read()

    root = ET.fromstring(kml_data)
    ns = {"kml": "http://www.opengis.net/kml/2.2"}

    records = []
    counter = 1

    for placemark in root.findall(".//kml:Placemark", ns):
        desc = placemark.findtext("kml:description", "", ns)

        square, sign = "", ""

        # ===== استخراج من Description (الطريقة القديمة) =====
        if desc:
            soup = BeautifulSoup(desc, "html.parser")
            tds = [td.get_text(strip=True) for td in soup.find_all("td")]
            for i in range(len(tds)):
                if tds[i] == "رقم المربع":
                    square = tds[i + 1]
                if tds[i] == "رقم الشاخص":
                    sign = tds[i + 1]

        # ===== استخراج من ExtendedData (الطريقة الجديدة) =====
        extended_data = placemark.find("kml:ExtendedData", ns)
        if extended_data is not None:
            for data in extended_data.findall("kml:Data", ns):
                name = data.get("name")
                value = data.findtext("kml:value", "", ns)
                if name == "square_number":
                    square = value
                elif name == "sign_number":
                    sign = value

        for poly in placemark.findall(".//kml:Polygon", ns):
            coords_text = poly.findtext(".//kml:coordinates", "", ns).strip()
            coords = []

            for c in coords_text.split():
                lon, lat, *_ = c.split(",")
                coords.append((float(lon), float(lat)))

            polygon_shape = Polygon(coords)
            
            # حساب المساحة بالمتر المربع
            area = calculate_area_in_sqm(polygon_shape)
            
            # حساب نقطة المنتصف
            centroid = polygon_shape.centroid
            center_coords = (centroid.x, centroid.y)

            records.append({
                "polygon_id": counter,
                "square_number": square,
                "sign_number": sign,
                "coordinates": coords,
                "polygon": polygon_shape,
                "Area": area,
                "Center": center_coords
            })
            counter += 1

    df = pd.DataFrame(records)
    spatial_index = STRtree(df["polygon"].tolist())
    return df, spatial_index


def load_polygons_from_excel(uploaded_excel):
    """تحميل زونات من Excel"""
    df = pd.read_excel(uploaded_excel)
    df["coordinates"] = df["coordinates"].apply(json.loads)
    df["polygon"] = df["coordinates"].apply(lambda c: Polygon(c))
    
    # إضافة Area و Center إذا لم تكن موجودة
    if "Area" not in df.columns:
        df["Area"] = df["polygon"].apply(calculate_area_in_sqm)
    if "Center" not in df.columns:
        df["Center"] = df["polygon"].apply(lambda p: (p.centroid.x, p.centroid.y))
    
    spatial_index = STRtree(df["polygon"].tolist())
    return df, spatial_index


def find_point(point, df, spatial_index):
    """إيجاد الزونات التي تحتوي نقطة"""
    results = []
    for idx in spatial_index.query(point):
        poly = df.iloc[idx]["polygon"]
        if poly.covers(point):
            row = df.iloc[idx]
            results.append({
                "polygon_id": int(row["polygon_id"]),
                "square_number": row["square_number"],
                "sign_number": row["sign_number"]
            })
    return results

def compare_zone_data(point_square, point_sign, zone_data):
    """مقارنة بيانات النقطة مع بيانات الزون"""
    if not zone_data:  # لا يوجد زونات
        return None, None, 4
    
    zone = zone_data[0]  # نأخذ أول زون (في حالة تعدد الزونات)
    zone_square = zone.get("square_number", "")
    zone_sign = zone.get("sign_number", "")
    
    cmp_sign = "T" if str(point_sign) == str(zone_sign) else "F"
    cmp_square = "T" if str(point_square) == str(zone_square) else "F"
    
    # تحديد CMP_Result
    if cmp_sign == "T":
        cmp_result = 1
    elif cmp_square == "T" and cmp_sign == "F":
        cmp_result = 2
    elif cmp_square == "F":
        cmp_result = 3
    else:
        cmp_result = 3
    
    return cmp_sign, cmp_square, cmp_result


def find_nearest_zone(point, df):
    """إيجاد أقرب زون بناء على نقطة المنتصف"""
    min_dist = float('inf')
    nearest_zone = None
    
    for _, row in df.iterrows():
        center = Point(row["Center"])
        dist = calculate_distance_in_meters(point, center)
        if dist < min_dist:
            min_dist = dist
            nearest_zone = {
                "polygon_id": int(row["polygon_id"]),
                "square_number": row["square_number"],
                "sign_number": row["sign_number"],
                "distance_meters": round(dist, 2)
            }
    
    return nearest_zone

def export_kml(df, fill_alpha):
    doc = Document()

    kml = doc.createElement("kml")
    kml.setAttribute("xmlns", "http://www.opengis.net/kml/2.2")
    doc.appendChild(kml)

    document = doc.createElement("Document")
    kml.appendChild(document)

    for _, row in df.iterrows():
        placemark = doc.createElement("Placemark")

        # ===== الاسم =====
        name = doc.createElement("name")
        name.appendChild(doc.createTextNode(f"Polygon {row['polygon_id']}"))
        placemark.appendChild(name)

        # ===== الوصف =====
        desc = doc.createElement("description")
        desc.appendChild(
            doc.createTextNode(
                f"Square: {row['square_number']} | Sign: {row['sign_number']} | Area: {row['Area']:.2f} m²"
            )
        )
        placemark.appendChild(desc)

        # ===== ExtendedData =====
        extended_data = doc.createElement("ExtendedData")
        
        for field, value in [("square_number", row['square_number']), 
                             ("sign_number", row['sign_number']),
                             ("Area_sqm", f"{row['Area']:.2f}")]:
            data = doc.createElement("Data")
            data.setAttribute("name", field)
            val = doc.createElement("value")
            val.appendChild(doc.createTextNode(str(value)))
            data.appendChild(val)
            extended_data.appendChild(data)
        
        placemark.appendChild(extended_data)

        # ===== النمط =====
        fill, line = random_kml_color(fill_alpha, "FF")

        style = doc.createElement("Style")

        line_style = doc.createElement("LineStyle")
        lc = doc.createElement("color")
        lc.appendChild(doc.createTextNode(line))
        lw = doc.createElement("width")
        lw.appendChild(doc.createTextNode("2"))
        line_style.appendChild(lc)
        line_style.appendChild(lw)

        poly_style = doc.createElement("PolyStyle")
        pc = doc.createElement("color")
        pc.appendChild(doc.createTextNode(fill))
        poly_style.appendChild(pc)

        style.appendChild(line_style)
        style.appendChild(poly_style)
        placemark.appendChild(style)

        # ===== Polygon =====
        polygon = doc.createElement("Polygon")

        outer = doc.createElement("outerBoundaryIs")
        ring = doc.createElement("LinearRing")
        coords = doc.createElement("coordinates")

        coord_text = " ".join(
            f"{lon},{lat},0"
            for lon, lat in row["polygon"].exterior.coords
        )

        coords.appendChild(doc.createTextNode(coord_text))
        ring.appendChild(coords)
        outer.appendChild(ring)
        polygon.appendChild(outer)

        placemark.appendChild(polygon)
        document.appendChild(placemark)

    return doc.toprettyxml(indent="  ")

# ======================================================
# الواجهة – اختيار مصدر الزونات
# ======================================================
st.subheader("🗂️ مصدر الزونات")

source = st.radio(
    "اختر طريقة إدخال الزونات:",
    ["KMZ / KML", "Excel"]
)

df_polygons = None
spatial_index = None


# ======================================================
# خيار Excel
# ======================================================
if source == "Excel":
    st.info("📄 قالب اكسل المطلوب:")
    st.code("""
polygon_id | square_number | sign_number | coordinates
1          | 29A           | 2/204       | [[lon,lat],[lon,lat],...]
""")

    uploaded_excel = st.file_uploader("📂 ارفع ملف Excel للزونات", type=["xlsx"])
    if uploaded_excel:
        df_polygons, spatial_index = load_polygons_from_excel(uploaded_excel)
        st.success(f"تم تحميل {len(df_polygons)} زون")


# ======================================================
# خيار KMZ / KML
# ======================================================
if source == "KMZ / KML":
    uploaded_file = st.file_uploader(
        "📂 ارفع ملف KMZ أو KML",
        type=["kmz", "kml"]
    )

    if uploaded_file:
        df_polygons, spatial_index = parse_kmz_or_kml(uploaded_file)
        st.success(f"تم تحميل {len(df_polygons)} زون")

        export_df = df_polygons.drop(columns=["polygon"]).copy()
        export_df["coordinates"] = export_df["coordinates"].apply(json.dumps)
        export_df["Center"] = export_df["Center"].apply(json.dumps)  # إضافة هذا السطر

        buffer = io.BytesIO()
        export_df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            "📥 تحميل الزونات كـ Excel",
            data=buffer,
            file_name="polygons.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ======================================================
# باقي الخطوات (مشتركة)
# ======================================================
if df_polygons is not None:

    st.subheader("📍 اختبار نقطة واحدة")
    coord_text = st.text_input("أدخل الإحداثيات كالتالي: (lat, lon). Ex: 21.41855, 39.88040")

    if st.button("اختبار النقطة"):
        lat, lon = map(float, coord_text.split(","))
        point = Point(lon, lat)
        st.json(find_point(point, df_polygons, spatial_index))
        
        
    st.subheader("📊 اختبار ملف نقاط Excel")
    excel_points = st.file_uploader("ارفع ملف Excel (location_type, id, lat, lon, square_number, sign_number)", type=["xlsx"])

    if excel_points:
        points_df = pd.read_excel(excel_points)
        
        # التأكد من وجود العواميد المطلوبة
        if "square_number" not in points_df.columns:
            points_df["square_number"] = ""
        if "sign_number" not in points_df.columns:
            points_df["sign_number"] = ""
        
        results = []

        for _, r in points_df.iterrows():
            p = Point(r["lon"], r["lat"])
            matches = find_point(p, df_polygons, spatial_index)
            
            # المقارنة
            cmp_sign, cmp_square, cmp_result = compare_zone_data(
                r.get("square_number", ""),
                r.get("sign_number", ""),
                matches
            )
            
            # إنشاء السطر مع الحفاظ على جميع البيانات الأصلية
            result_row = r.to_dict()  # نحتفظ بجميع البيانات الأصلية
            
            # إضافة النتائج الجديدة
            result_row.update({
                "polygons_count": len(matches),
                "result": json.dumps(matches, ensure_ascii=False),
                "CMP_square": cmp_square if cmp_square else "",
                "CMP_sign": cmp_sign if cmp_sign else "",
                "CMP_Result": cmp_result
            })
            
            results.append(result_row)

        out_df = pd.DataFrame(results)
        st.dataframe(out_df)

        buf = io.BytesIO()
        out_df.to_excel(buf, index=False)
        buf.seek(0)

        st.download_button(
            "📥 تحميل نتائج النقاط",
            data=buf,
            file_name="points_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.subheader("🔍 إيجاد أقرب زون للنقاط (3 و 4)")
    st.info("هذا القسم يعمل على ملف النتائج السابق ويجد أقرب زون للنقاط ذات CMP_Result = 3 أو 4")
    
    excel_results = st.file_uploader("ارفع ملف نتائج النقاط (points_result.xlsx)", type=["xlsx"], key="nearest_zone_file")

    if excel_results:
        results_df = pd.read_excel(excel_results)
        
        if "CMP_Result" not in results_df.columns:
            st.error("الملف المرفوع لا يحتوي على عامود CMP_Result. تأكد من رفع ملف النتائج الصحيح.")
        else:
            if st.button("🔎 ابحث عن أقرب زون"):
                updated_results = []
                
                progress_bar = st.progress(0)
                total = len(results_df)
                
                for idx, r in results_df.iterrows():
                    result_row = r.to_dict()
                    
                    # فقط للنقاط ذات CMP_Result = 3 أو 4
                    if r["CMP_Result"] in [3, 4]:
                        p = Point(r["lon"], r["lat"])
                        nearest = find_nearest_zone(p, df_polygons)
                        result_row["nearest_zone"] = json.dumps(nearest, ensure_ascii=False) if nearest else ""
                    else:
                        result_row["nearest_zone"] = ""
                    
                    updated_results.append(result_row)
                    progress_bar.progress((idx + 1) / total)
                
                final_df = pd.DataFrame(updated_results)
                st.success(f"تم معالجة {len(final_df)} نقطة")
                st.dataframe(final_df)
                
                buf = io.BytesIO()
                final_df.to_excel(buf, index=False)
                buf.seek(0)
                
                st.download_button(
                    "📥 تحميل النتائج مع أقرب زون",
                    data=buf,
                    file_name="points_with_nearest_zone.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    st.subheader("🧩 تصدير KML")
    alpha = st.slider("شفافية داخل الزون", 0, 255, 85)

    if st.button("تحميل KML"):
        st.download_button(
            "تحميل zones.kml",
            export_kml(df_polygons, f"{alpha:02x}"),
            file_name="zones.kml"
        )


# =============================================================================
# import io
# import streamlit as st
# import zipfile
# import xml.etree.ElementTree as ET
# import pandas as pd
# import json
# import random
# from shapely.geometry import Polygon, Point
# from shapely.strtree import STRtree
# from bs4 import BeautifulSoup
# from datetime import datetime
# from xml.dom.minidom import Document
# 
# # =========================
# # إعداد الصفحة
# # =========================
# st.set_page_config(
#     page_title="أداة فحص الزونات (KMZ)",
#     layout="wide"
# )
# 
# st.title("🗺️ أداة استخراج وفحص الزونات (KMZ)")
# 
# # =========================
# # دوال مساعدة
# # =========================
# 
# def random_kml_color(fill_alpha="55", line_alpha="FF"):
#     r = random.randint(0, 255)
#     g = random.randint(0, 255)
#     b = random.randint(0, 255)
#     return (
#         f"{fill_alpha}{b:02x}{g:02x}{r:02x}",
#         f"{line_alpha}{b:02x}{g:02x}{r:02x}"
#     )
# 
# 
# def parse_kmz(uploaded_file):
#     with zipfile.ZipFile(uploaded_file, "r") as kmz:
#         kml_name = [f for f in kmz.namelist() if f.endswith(".kml")][0]
#         kml_data = kmz.read(kml_name)
# 
#     root = ET.fromstring(kml_data)
#     ns = {"kml": "http://www.opengis.net/kml/2.2"}
# 
#     records = []
#     counter = 1
# 
#     for placemark in root.findall(".//kml:Placemark", ns):
#         desc = placemark.findtext("kml:description", "", ns)
# 
#         square = ""
#         sign = ""
# 
#         if desc:
#             soup = BeautifulSoup(desc, "html.parser")
#             tds = [td.get_text(strip=True) for td in soup.find_all("td")]
#             for i in range(len(tds)):
#                 if tds[i] == "رقم المربع":
#                     square = tds[i + 1]
#                 if tds[i] == "رقم الشاخص":
#                     sign = tds[i + 1]
# 
#         for poly in placemark.findall(".//kml:Polygon", ns):
#             coords_text = poly.findtext(".//kml:coordinates", "", ns).strip()
#             coords = []
#             for c in coords_text.split():
#                 lon, lat, *_ = c.split(",")
#                 coords.append((float(lon), float(lat)))
# 
#             records.append({
#                 "polygon_id": counter,
#                 "square_number": square,
#                 "sign_number": sign,
#                 "polygon": Polygon(coords),
#                 "coordinates": coords
#             })
#             counter += 1
# 
#     df = pd.DataFrame(records)
#     tree = STRtree(df["polygon"].tolist())
#     return df, tree
# 
# 
# def find_point(point, df, tree):
#     results = []
# 
#     candidate_indexes = tree.query(point)
# 
#     for idx in candidate_indexes:
#         poly = df.iloc[idx]["polygon"]
# 
#         if poly.covers(point):
#             row = df.iloc[idx]
#             results.append({
#                 "polygon_id": int(row["polygon_id"]),
#                 "square_number": row["square_number"],
#                 "sign_number": row["sign_number"]
#             })
# 
#     return results
# 
# 
# 
# def export_kml(df, fill_alpha):
#     doc = Document()
#     kml = doc.createElement("kml")
#     kml.setAttribute("xmlns", "http://www.opengis.net/kml/2.2")
#     doc.appendChild(kml)
# 
#     document = doc.createElement("Document")
#     kml.appendChild(document)
# 
#     for _, row in df.iterrows():
#         placemark = doc.createElement("Placemark")
# 
#         name = doc.createElement("name")
#         name.appendChild(doc.createTextNode(f"Polygon {row['polygon_id']}"))
#         placemark.appendChild(name)
# 
#         fill, line = random_kml_color(fill_alpha, "FF")
# 
#         style = doc.createElement("Style")
# 
#         ls = doc.createElement("LineStyle")
#         lc = doc.createElement("color")
#         lc.appendChild(doc.createTextNode(line))
#         lw = doc.createElement("width")
#         lw.appendChild(doc.createTextNode("2"))
#         ls.appendChild(lc)
#         ls.appendChild(lw)
# 
#         ps = doc.createElement("PolyStyle")
#         pc = doc.createElement("color")
#         pc.appendChild(doc.createTextNode(fill))
#         ps.appendChild(pc)
# 
#         style.appendChild(ls)
#         style.appendChild(ps)
#         placemark.appendChild(style)
# 
#         polygon = doc.createElement("Polygon")
#         outer = doc.createElement("outerBoundaryIs")
#         ring = doc.createElement("LinearRing")
#         coords = doc.createElement("coordinates")
# 
#         coord_text = " ".join(
#             f"{lon},{lat},0" for lon, lat in row["polygon"].exterior.coords
#         )
# 
#         coords.appendChild(doc.createTextNode(coord_text))
#         ring.appendChild(coords)
#         outer.appendChild(ring)
#         polygon.appendChild(outer)
#         placemark.appendChild(polygon)
# 
#         document.appendChild(placemark)
# 
#     return doc.toprettyxml()
# 
# def load_polygons_from_excel(uploaded_excel):
#     df = pd.read_excel(uploaded_excel)
#     df["coordinates"] = df["coordinates"].apply(json.loads)
#     df["polygon"] = df["coordinates"].apply(lambda c: Polygon(c))
#     spatial_index = STRtree(df["polygon"].tolist())
#     return df, spatial_index
# 
# 
# # =========================
# # الواجهة
# # =========================
# 
# source = st.radio(
#     "اختر مصدر الزونات",
#     ["KMZ / KML", "Excel"]
# )
# 
# if source == "Excel":
#     uploaded_excel = st.file_uploader("📂 ارفع ملف Excel للزونات", type=["xlsx"])
#     if uploaded_excel:
#         df_polygons, spatial_index = load_polygons_from_excel(uploaded_excel)
# 
# 
# uploaded_kmz = st.file_uploader("📂 ارفع ملف KMZ", type=["kmz"])
# 
# 
# if source == "KMZ / KML":
#     if uploaded_kmz:
#         df_polygons, spatial_index = parse_kmz(uploaded_kmz)
#         st.success(f"تم تحميل {len(df_polygons)} زون")
#         export_df = df_polygons.copy()
#         export_df["coordinates"] = export_df["coordinates"].apply(json.dumps)
#         export_df.drop(columns=["polygon"], inplace=True)
#     
#             # ===== تصدير Excel (زر تحميل) =====
#         excel_buffer = io.BytesIO()
#         export_df.to_excel(excel_buffer, index=False)
#         excel_buffer.seek(0)
#         
#         st.download_button(
#             label="📥 تحميل بيانات الزونات (Excel)",
#             data=excel_buffer,
#             file_name="polygons.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
# 
#     
#     
#     st.subheader("📍 اختبار نقطة واحدة")
# 
#     coord_text = st.text_input("أدخل الإحداثيات كالتالي: (lat, lon). Ex: 21.41855, 39.88040")
#     
#     if st.button("اختبار النقطة"):
#         try:
#             lat, lon = map(float, coord_text.split(","))
#             point = Point(lon, lat)
#             matches = find_point(point, df_polygons, spatial_index)
#             st.write(f"عدد الزونات: {len(matches)}")     
#             st.json(matches)
#         except Exception:
#             st.error("صيغة الإحداثيات غير صحيحة. استخدم: lat, lon")
# 
#     st.subheader("📊 اختبار ملف Excel")
#     excel_file = st.file_uploader("ارفع ملف Excel (id, lat, lon)", type=["xlsx"])
# 
#     if excel_file:
#         points_df = pd.read_excel(excel_file)
#         results = []
# 
#         for _, r in points_df.iterrows():
#             p = Point(r["lon"], r["lat"])
#             matches = find_point(p, df_polygons, spatial_index)
#             results.append({
#                 "id": r["id"],
#                 "lat": r["lat"],
#                 "lon": r["lon"],
#                 "عدد_الزونات": len(matches),
#                 "النتيجة": json.dumps(matches, ensure_ascii=False)
#             })
# 
#         out_df = pd.DataFrame(results)
#         st.dataframe(out_df)
# 
#         excel_buffer = io.BytesIO()
#         out_df.to_excel(excel_buffer, index=False)
#         excel_buffer.seek(0)
#         
#         st.download_button(
#             label="📥 تحميل نتيجة اختبار النقاط (Excel)",
#             data=excel_buffer,
#             file_name="Result_points.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
# 
# 
#     st.subheader("🧩 تصدير KML")
#     alpha = st.slider("شفافية داخل الزون", 0, 255, 85)
# 
#     if st.button("تحميل KML"):
#         kml_data = export_kml(df_polygons, f"{alpha:02x}")
#         st.download_button(
#             "تحميل zones.kml",
#             kml_data,
#             file_name="zones.kml"
#         )
# 
# =============================================================================
