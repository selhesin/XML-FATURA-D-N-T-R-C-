import streamlit as st
import pandas as pd
import io
import xml.etree.ElementTree as ET
import time

def efatura_mal_iceriklerini_oku(uploaded_files):
    """
    UBL-TR E-Fatura XML dosyalarından satıcı, alıcı, VKN, mal/hizmet satırları (InvoiceLine) 
    ve tutar detaylarını eksiksiz ayrıştırır.
    """
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
            
            # 1. Genel Fatura Başlık ve Taraf Bilgileri
            for elem in root.iter():
                tag = elem.tag.split('}')[-1].lower()
                text = elem.text.strip() if elem.text else ""
                
                if tag in ["id", "faturano"] and fatura_no == "Bilinmiyor" and text:
                    fatura_no = text
                elif tag in ["issuedate", "tarih"] and not fatura_tarihi and text:
                    fatura_tarihi = text

            # Satıcı ve Alıcı Ayrımı (UBL-TR Yapısı)
            for parent in root.iter():
                ptag = parent.tag.split('}')[-1].lower()
                if "accountingsupplierparty" in ptag: # Satıcı
                    for child in parent.iter():
                        ctag = child.tag.split('}')[-1].lower()
                        cval = child.text.strip() if child.text else ""
                        if ctag in ["name", "partyname"] and satici_unvan == "Bilinmiyor" and cval:
                            satici_unvan = cval
                        elif ctag in ["companyid", "vkn", "tckn"] and not satici_vkn and cval:
                            satici_vkn = cval
                elif "accountingcustomerparty" in ptag: # Alıcı
                    for child in parent.iter():
                        ctag = child.tag.split('}')[-1].lower()
                        cval = child.text.strip() if child.text else ""
                        if ctag in ["name", "partyname"] and alici_unvan == "Bilinmiyor" and cval:
                            alici_unvan = cval
                        elif ctag in ["companyid", "vkn", "tckn"] and not alici_vkn and cval:
                            alici_vkn = cval

            # Eğer genel etiketlerde unvan bulunamadıysa genel arama yap
            if satici_unvan == "Bilinmiyor":
                for elem in root.iter():
                    tag = elem.tag.split('}')[-1].lower()
                    if tag in ["partyname", "name", "unvan"] and elem.text:
                        satici_unvan = elem.text.strip()
                        break

            # 2. Fatura Mal / Hizmet Kalemleri (InvoiceLine)
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

def render_ft_module():
    """
    Streamlit arayüzünde mal/hizmet kalemleri sekmesini ve indirme butonunu çizen ana fonksiyon.
    """
    st.subheader("📦 E-Fatura Mal / Hizmet & Taraf Detay Analizi (`ft.py`)")
    st.markdown("E-Fatura XML dosyalarınızı yükleyerek satıcı, alıcı, VKN ve ürün/hizmet kalemlerini tek tabloda listeleyebilir ve Excel olarak indirebilirsiniz.")
    
    efatura_kalem_files = st.file_uploader(
        "E-Fatura XML Dosyalarını Seçin", 
        type=["xml"], 
        accept_multiple_files=True,
        key="ft_efatura_uploader"
    )
    
    if st.button("🚀 Faturalardaki Detayları ve Kalemleri Ayrıştır", type="primary", use_container_width=True):
        if efatura_kalem_files:
            with st.spinner("Faturalar taranıyor, satıcı/alıcı ve kalemler çıkartılıyor..."):
                df_kalemler = efatura_mal_iceriklerini_oku(efatura_kalem_files)
                st.session_state["ft_efatura_kalemler_df"] = df_kalemler
                time.sleep(0.3)
            st.success("Tüm fatura detayları ve kalemleri başarıyla listelendi kanka!")
            st.rerun()
        else:
            st.warning("Lütfen en az bir E-Fatura XML dosyası seçin kanka.")

    if "ft_efatura_kalemler_df" in st.session_state and not st.session_state["ft_efatura_kalemler_df"].empty:
        st.markdown("---")
        st.markdown("### 📋 Fatura Kalemleri ve Taraf Listesi")
        
        df_gosterim = st.session_state["ft_efatura_kalemler_df"].copy()
        for col in ["Birim Fiyat (TL)", "Matrah (TL)", "KDV Tutarı (TL)", "Toplam (TL)"]:
            df_gosterim[col] = df_gosterim[col].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            
        st.dataframe(df_gosterim, use_container_width=True, hide_index=True)
        
        output_kalemler = io.BytesIO()
        with pd.ExcelWriter(output_kalemler, engine='xlsxwriter') as writer:
            st.session_state["ft_efatura_kalemler_df"].to_excel(writer, sheet_name='Fatura_Detay_Kalemleri', index=False)
            
        st.download_button(
            label="📥 Bu Detaylı Listeyi Excel Olarak İndir (.xlsx)", 
            data=output_kalemler.getvalue(), 
            file_name="E_Fatura_Detayli_Mal_Hizmet_Listesi.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            key="dl_ft_efatura_kalemler"
        )
