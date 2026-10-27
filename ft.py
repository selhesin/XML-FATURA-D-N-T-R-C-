import streamlit as st
import pandas as pd
import io
import xml.etree.ElementTree as ET
import time

# Sayfa Ayarları
st.set_page_config(
    page_title="E-Fatura Alış KDV Listesi & Analiz Portalı",
    page_icon="🧾",
    layout="wide"
)

# Başlık ve Karşılama
st.title("🧾 E-Fatura Alış KDV Listesi Otomasyonu")
st.markdown("Alış faturası XML dosyalarınızı yükleyin; sistem satıcı, VKN, mal/hizmet kalemleri, KDV ve tam fatura numarasını resmi **İndirilecek KDV Listesi** formatına dönüştürüp Excel versin.")
st.markdown("---")

# Oturum Durumu Başlatma
if "ft_efatura_kalemler_df" not in st.session_state:
    st.session_state["ft_efatura_kalemler_df"] = pd.DataFrame()

def efatura_mal_iceriklerini_oku(uploaded_files):
    tum_satirlar = []
    
    for uf in uploaded_files:
        try:
            tree = ET.parse(uf)
            root = tree.getroot()
            
            fatura_no = "Bilinmiyor"
            fatura_tarihi = ""
            satici_unvan = "Bilinmiyor"
            satici_vkn = ""
            
            # 1. Genel Fatura Bilgileri (Tam Fatura No ve Tarih)
            for elem in root.iter():
                tag = elem.tag.split('}')[-1].lower()
                text = elem.text.strip() if elem.text else ""
                
                if tag in ["id", "faturano"] and fatura_no == "Bilinmiyor" and text:
                    fatura_no = text  # Seri ayrımı olmadan direkt tam fatura no
                elif tag in ["issuedate", "tarih"] and not fatura_tarihi and text:
                    fatura_tarihi = text

            # Tarih formatını YYYY-AA-GG -> GG.AA.YYYY çevirme (varsa)
            try:
                if "-" in fatura_tarihi:
                    parts = fatura_tarihi.split("-")
                    if len(parts) == 3:
                        fatura_tarihi = f"{parts[2]}.{parts[1]}.{parts[0]}"
            except:
                pass

            # Dönem tespiti (Örn: 2023/05)
            kdv_donemi = ""
            try:
                if "-" in root.find('.//{*}IssueDate').text:
                    p = root.find('.//{*}IssueDate').text.split('-')
                    kdv_donemi = f"{p[0]}/{p[1]}"
            except:
                kdv_donemi = "2023/05"

            # 2. Satıcı Bilgileri (AccountingSupplierParty)
            for parent in root.iter():
                ptag = parent.tag.split('}')[-1].lower()
                if "accountingsupplierparty" in ptag:
                    for child in parent.iter():
                        ctag = child.tag.split('}')[-1].lower()
                        cval = child.text.strip() if child.text else ""
                        if ctag in ["name", "partyname"] and satici_unvan == "Bilinmiyor" and cval:
                            satici_unvan = cval
                        elif ctag in ["companyid", "vkn", "tckn"] and not satici_vkn and cval:
                            satici_vkn = cval

            if satici_unvan == "Bilinmiyor":
                for elem in root.iter():
                    tag = elem.tag.split('}')[-1].lower()
                    if tag in ["partyname", "name", "unvan"] and elem.text:
                        satici_unvan = elem.text.strip()
                        break

            # 3. Mal ve Hizmet Kalemleri (InvoiceLine)
            for line in root.iter():
                tag = line.tag.split('}')[-1].lower()
                if tag in ["invoiceline", "line"]:
                    mal_adi = "Tanımsız Mal / Hizmet"
                    miktar = 1.0
                    tutar = 0.0
                    kdv_orani = 18.0
                    
                    for child in line.iter():
                        ctag = child.tag.split('}')[-1].lower()
                        cval = child.text.strip() if child.text else ""
                        
                        if ctag in ["name", "itemname", "description"] and cval and mal_adi == "Tanımsız Mal / Hizmet":
                            mal_adi = cval
                        elif ctag in ["invoicedquantity", "quantity"]:
                            try: miktar = float(cval)
                            except: pass
                        elif ctag in ["lineextensionamount", "amount"]:
                            try: tutar = float(cval)
                            except: pass
                        elif ctag in ["percent"] and cval in ["1", "8", "10", "18", "20"]:
                            try: kdv_orani = float(cval)
                            except: pass

                    kdv_tutari = tutar * (kdv_orani / 100.0)

                    tum_satirlar.append({
                        "Alış Faturasının Tarihi": fatura_tarihi,
                        "Fatura Numarası": fatura_no,
                        "Satıcının Adı-Soyadı / Ünvanı": satici_unvan,
                        "Satıcının VKN / TCKN": satici_vkn,
                        "Alınan Mal ve/veya Hizmetin Cinsi": mal_adi,
                        "Alınan Mal ve/veya Hizmetin Miktarı": miktar,
                        "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı": tutar,
                        "KDV'si": kdv_tutari,
                        "Tevkifatlı Faturanın Tevkifata Tabi Olmayan Ve Bu Dönemde İndirilen Kdv Tutarı": 0.0,
                        "2 Nolu Beyannamede Ödenen Kdv Tutarı": 0.0,
                        "Toplam İndirilen KDV Tutarı": kdv_tutari,
                        "GGB Tescil No'su (Alış İthalat İse)": "",
                        "Belgenin İndirim Hakkının Kullanıldığı KDV Dönemi": kdv_donemi
                    })
        except Exception as e:
            st.error(f"Dosya okuma hatası ({uf.name}): {e}")
            
    # Sıra No ekleyelim
    df = pd.DataFrame(tum_satirlar)
    if not df.empty:
        df.insert(0, "Sıra No", range(1, len(df) + 1))
    return df

# --- ARAYÜZ KISMI ---
uploaded_xmls = st.file_uploader(
    "Alış Faturası XML Dosyalarını Seçin", 
    type=["xml"], 
    accept_multiple_files=True,
    key="ft_efatura_uploader"
)

col_islem, col_temizle = st.columns([3, 1])

with col_islem:
    islem_basildi = st.button("🚀 Faturaları İşle ve İndirilecek KDV Listesi Oluştur", type="primary", use_container_width=True)

with col_temizle:
    temizle_basildi = st.button("🗑️ Faturaları Temizle / Sıfırla", type="secondary", use_container_width=True)

if temizle_basildi:
    st.session_state["ft_efatura_kalemler_df"] = pd.DataFrame()
    st.success("Tüm fatura verileri temizlendi kanka!")
    st.rerun()

if islem_basildi:
    if uploaded_xmls:
        with st.spinner("XML'ler taranıyor, resmi İndirilecek KDV Listesi formatına dönüştürülüyor..."):
            df_sonuc = efatura_mal_iceriklerini_oku(uploaded_xmls)
            st.session_state["ft_efatura_kalemler_df"] = df_sonuc
            time.sleep(0.3)
        st.success("İşlem tamam kanka!")
        st.rerun()
    else:
        st.warning("Lütfen önce XML dosyası seçin kanka.")

if not st.session_state["ft_efatura_kalemler_df"].empty:
    st.markdown("---")
    st.subheader("📋 Resmi İndirilecek KDV Listesi Tablosu")
    
    df_goster = st.session_state["ft_efatura_kalemler_df"].copy()
    for col in ["Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı", "KDV'si", "Toplam İndirilen KDV Tutarı"]:
        df_goster[col] = df_goster[col].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
    st.dataframe(df_goster, use_container_width=True, hide_index=True)
    
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        st.session_state["ft_efatura_kalemler_df"].to_excel(writer, sheet_name='İndirilecek KDV Listesi', index=False)
        
    st.download_button(
        label="📥 İndirilecek KDV Listesini Excel Olarak İndir (.xlsx)", 
        data=output_excel.getvalue(), 
        file_name="Indirilecek_KDV_Listesi.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        key="dl_excel_btn"
    )
