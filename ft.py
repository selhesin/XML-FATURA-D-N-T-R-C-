import streamlit as st
import pandas as pd
import io
import xml.etree.ElementTree as ET
import time

# Sayfa Ayarları
st.set_page_config(
    page_title="E-Fatura Alış Kalemleri & KDV Analiz Portalı",
    page_icon="🧾",
    layout="wide"
)

# Başlık ve Karşılama
st.title("🧾 E-Fatura Alış / Satış Detay ve KDV Analiz Portalı")
st.markdown("E-Fatura XML dosyalarınızı yükleyin; sistem satıcı, alıcı, VKN, mal/hizmet kalemleri, KDV ve toplam tutarları otomatik ayrıştırıp size Excel versin.")
st.markdown("---")

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
            alici_unvan = "Bilinmiyor"
            alici_vkn = ""
            
            # Genel Fatura Bilgileri
            for elem in root.iter():
                tag = elem.tag.split('}')[-1].lower()
                text = elem.text.strip() if elem.text else ""
                
                if tag in ["id", "faturano"] and fatura_no == "Bilinmiyor" and text:
                    fatura_no = text
                elif tag in ["issuedate", "tarih"] and not fatura_tarihi and text:
                    fatura_tarihi = text

            # Satıcı ve Alıcı Ayrımı
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
                elif "accountingcustomerparty" in ptag:
                    for child in parent.iter():
                        ctag = child.tag.split('}')[-1].lower()
                        cval = child.text.strip() if child.text else ""
                        if ctag in ["name", "partyname"] and alici_unvan == "Bilinmiyor" and cval:
                            alici_unvan = cval
                        elif ctag in ["companyid", "vkn", "tckn"] and not alici_vkn and cval:
                            alici_vkn = cval

            if satici_unvan == "Bilinmiyor":
                for elem in root.iter():
                    tag = elem.tag.split('}')[-1].lower()
                    if tag in ["partyname", "name", "unvan"] and elem.text:
                        satici_unvan = elem.text.strip()
                        break

            # Mal ve Hizmet Kalemleri (InvoiceLine)
            for line in root.iter():
                tag = line.tag.split('}')[-1].lower()
                if tag in ["invoiceline", "line"]:
                    mal_adi = "Tanımsız Mal / Hizmet"
                    miktar = 1.0
                    birim_fiyat = 0.0
                    tutar = 0.0
                    kdv_orani = 20.0
                    
                    for child in line.iter():
                        ctag = child.tag.split('}')[-1].lower()
                        cval = child.text.strip() if child.text else ""
                        
                        if ctag in ["name", "itemname", "description"] and cval and mal_adi == "Tanımsız Mal / Hizmet":
                            mal_adi = cval
                        elif ctag in ["invoicedquantity", "quantity"]:
                            try: miktar = float(cval)
                            except: pass
                        elif ctag in ["priceamount", "price"]:
                            try: birim_fiyat = float(cval)
                            except: pass
                        elif ctag in ["lineextensionamount", "amount"]:
                            try: tutar = float(cval)
                            except: pass
                        elif ctag in ["percent"] and cval in ["1", "8", "10", "18", "20"]:
                            try: kdv_orani = float(cval)
                            except: pass
                    
                    if birim_fiyat == 0.0 and miktar > 0 and tutar > 0:
                        birim_fiyat = tutar / miktar
                    elif tutar == 0.0 and miktar > 0 and birim_fiyat > 0:
                        tutar = miktar * birim_fiyat

                    kdv_tutari = tutar * (kdv_orani / 100.0)
                    toplam_tutar = tutar + kdv_tutari

                    tum_satirlar.append({
                        "Fatura No": fatura_no,
                        "Tarih": fatura_tarihi,
                        "Satıcı Unvanı": satici_unvan,
                        "Satıcı VKN/TCKN": satici_vkn,
                        "Alıcı Unvanı": alici_unvan,
                        "Alıcı VKN/TCKN": alici_vkn,
                        "Mal / Hizmet Açıklaması": mal_adi,
                        "Miktar": miktar,
                        "Birim Fiyat (TL)": birim_fiyat,
                        "Matrah (TL)": tutar,
                        "KDV Oranı (%)": kdv_orani,
                        "KDV Tutarı (TL)": kdv_tutari,
                        "Toplam (TL)": toplam_tutar
                    })
        except Exception as e:
            st.error(f"Dosya okuma hatası ({uf.name}): {e}")
            
    return pd.DataFrame(tum_satirlar)

# --- ARAYÜZ KISMI ---
uploaded_xmls = st.file_uploader(
    "Alış Faturası XML Dosyalarını Seçin", 
    type=["xml"], 
    accept_multiple_files=True,
    key="main_efatura_uploader"
)

if st.button("🚀 Faturaları İşle ve Tabloyu Oluştur", type="primary", use_container_width=True):
    if uploaded_xmls:
        with st.spinner("XML'ler taranıyor, satıcı/alıcı ve kalemler ayrıştırılıyor..."):
            df_sonuc = efatura_mal_iceriklerini_oku(uploaded_xmls)
            st.session_state["islenen_fatura_df"] = df_sonuc
            time.sleep(0.3)
        st.success("İşlem tamam kanka!")
        st.rerun()
    else:
        st.warning("Lütfen önce XML dosyası seçin kanka.")

if "islenen_fatura_df" in st.session_state and not st.session_state["islenen_fatura_df"].empty:
    st.markdown("---")
    st.subheader("📊 Ayrıştırılan Fatura Kalemleri")
    
    df_goster = st.session_state["islenen_fatura_df"].copy()
    for col in ["Birim Fiyat (TL)", "Matrah (TL)", "KDV Tutarı (TL)", "Toplam (TL)"]:
        df_goster[col] = df_goster[col].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
    st.dataframe(df_goster, use_container_width=True, hide_index=True)
    
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        st.session_state["islenen_fatura_df"].to_excel(writer, sheet_name='E_Fatura_Analizi', index=False)
        
    st.download_button(
        label="📥 Bu Tabloyu Excel Olarak İndir (.xlsx)", 
        data=output_excel.getvalue(), 
        file_name="EFatura_Detayli_Analiz_Raporu.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        key="dl_excel_btn"
    )
