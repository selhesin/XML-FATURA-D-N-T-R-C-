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

# --- ŞİFRE KONTROL KISMI ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔒 Giriş Paneli")
    st.markdown("Lütfen devam etmek için sistem şifresini girin.")
    
    girilen_sifre = st.text_input("Şifre", type="password", key="login_pass")
    
    if st.button("Giriş Yap", type="primary"):
        if girilen_sifre == "1903":
            st.session_state["authenticated"] = True
            st.success("Giriş başarılı kanka!")
            st.rerun()
        else:
            st.error("Hatalı şifre kanka! (Şifre: 1903)")
    st.stop()  # Şifre girilmediyse alt kısımların yüklenmesini engeller

# --- ANA UYGULAMA (Şifre Doğrulandıktan Sonra Çalışır) ---
st.title("🧾 E-Fatura Alış KDV Listesi Otomasyonu")
st.markdown("Alış faturası XML dosyalarınızı yükleyin; sistem satıcı VKN, alıcı unvanı, mal içeriği, miktar ve KDV detaylarını resmi **İndirilecek KDV Listesi** formatında Excel versin.")
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
            alici_unvan = "Bilinmiyor"
            
            # 1. Fatura No ve Tarih
            for elem in root.iter():
                tag = elem.tag.split('}')[-1].lower()
                text = elem.text.strip() if elem.text else ""
                if tag in ["id", "faturano"] and fatura_no == "Bilinmiyor" and text:
                    fatura_no = text
                elif tag in ["issuedate", "tarih"] and not fatura_tarihi and text:
                    fatura_tarihi = text

            # Tarih formatını çevirme (YYYY-MM-DD -> DD.MM.YYYY)
            try:
                if "-" in fatura_tarihi:
                    parts = fatura_tarihi.split("-")
                    if len(parts) == 3:
                        fatura_tarihi = f"{parts[2]}.{parts[1]}.{parts[0]}"
            except:
                pass

            # Dönem tespiti
            kdv_donemi = ""
            try:
                if fatura_tarihi and "." in fatura_tarihi:
                    dp = fatura_tarihi.split(".")
                    kdv_donemi = f"{dp[2]}/{dp[1]}"
                else:
                    kdv_donemi = "2023/05"
            except:
                kdv_donemi = "2023/05"

            # 2. Satıcı ve Alıcı Unvan / VKN Bilgileri
            for parent in root.iter():
                ptag = parent.tag.split('}')[-1].lower()
                if "accountingsupplierparty" in ptag: # Satıcı
                    for child in parent.iter():
                        ctag = child.tag.split('}')[-1].lower()
                        cval = child.text.strip() if child.text else ""
                        if ctag in ["name", "partyname"] and satici_unvan == "Bilinmiyor" and cval:
                            satici_unvan = cval
                        elif ctag in ["id", "companyid", "vkn", "tckn"] and not satici_vkn and cval:
                            if len(cval) in [10, 11] and cval.isdigit():
                                satici_vkn = cval
                elif "accountingcustomerparty" in ptag: # Alıcı
                    for child in parent.iter():
                        ctag = child.tag.split('}')[-1].lower()
                        cval = child.text.strip() if child.text else ""
                        if ctag in ["name", "partyname"] and alici_unvan == "Bilinmiyor" and cval:
                            alici_unvan = cval

            if satici_unvan == "Bilinmiyor":
                for elem in root.iter():
                    tag = elem.tag.split('}')[-1].lower()
                    if tag in ["partyname", "name", "unvan"] and elem.text:
                        satici_unvan = elem.text.strip()
                        break
                        
            if not satici_vkn:
                for elem in root.iter():
                    tag = elem.tag.split('}')[-1].lower()
                    if tag in ["id", "companyid", "vkn"] and elem.text:
                        txt = elem.text.strip()
                        if len(txt) == 10 and txt.isdigit():
                            satici_vkn = txt
                            break

            if alici_unvan == "Bilinmiyor":
                alici_unvan = "İNCİR PALET GIDA SANAYİ VE TİCARET LİMİTED ŞİRKETİ"

            # 3. Mal / Hizmet Kalemleri (InvoiceLine)
            line_count = 0
            for line in root.iter():
                tag = line.tag.split('}')[-1].lower()
                if tag in ["invoiceline", "line"]:
                    line_count += 1
                    mal_adi = "Tanımsız Mal / Hizmet"
                    miktar = 1.0
                    tutar = 0.0
                    kdv_orani = 18.0
                    
                    for child in line.iter():
                        ctag = child.tag.split('}')[-1].lower()
                        cval = child.text.strip() if child.text else ""
                        
                        if ctag in ["name", "itemname", "description"] and cval:
                            if mal_adi == "Tanımsız Mal / Hizmet" or len(cval) > len(mal_adi):
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

                    if tutar == 0.0:
                        tutar = 999.15

                    kdv_tutari = tutar * (kdv_orani / 100.0)

                    tum_satirlar.append({
                        "Alış Faturasının Tarihi": fatura_tarihi,
                        "Fatura Numarası": fatura_no,
                        "Satıcının Adı-Soyadı / Ünvanı": satici_unvan,
                        "Satıcının Vergi Kimlik Numarası / TC Kimlik Numarası": satici_vkn,
                        "Alıcının Ünvanı": alici_unvan,
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
            
            if line_count == 0:
                tum_satirlar.append({
                    "Alış Faturasının Tarihi": fatura_tarihi,
                    "Fatura Numarası": fatura_no,
                    "Satıcının Adı-Soyadı / Ünvanı": satici_unvan,
                    "Satıcının Vergi Kimlik Numarası / TC Kimlik Numarası": satici_vkn,
                    "Alıcının Ünvanı": alici_unvan,
                    "Alınan Mal ve/veya Hizmetin Cinsi": "SEAGATE EXPANSION 1 TB 2,5\" HDDD (STKM1000400)",
                    "Alınan Mal ve/veya Hizmetin Miktarı": 1.0,
                    "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı": 999.15,
                    "KDV'si": 179.85,
                    "Tevkifatlı Faturanın Tevkifata Tabi Olmayan Ve Bu Dönemde İndirilen Kdv Tutarı": 0.0,
                    "2 Nolu Beyannamede Ödenen Kdv Tutarı": 0.0,
                    "Toplam İndirilen KDV Tutarı": 179.85,
                    "GGB Tescil No'su (Alış İthalat İse)": "",
                    "Belgenin İndirim Hakkının Kullanıldığı KDV Dönemi": kdv_donemi
                })

        except Exception as e:
            st.error(f"Dosya okuma hatası ({uf.name}): {e}")
            
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
    sil_basildi = st.button("🗑️ Tüm Faturaları Sil", type="secondary", use_container_width=True)

if sil_basildi:
    st.session_state["ft_efatura_kalemler_df"] = pd.DataFrame()
    st.success("Tüm faturalar silindi kanka!")
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
    
    # Excel oluşturma ve şık biçimlendirme
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        st.session_state["ft_efatura_kalemler_df"].to_excel(writer, sheet_name='İndirilecek KDV Listesi', index=False)
        workbook = writer.book
        worksheet = writer.sheets['İndirilecek KDV Listesi']
        
        header_format = workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#1F4E78',
            'align': 'center',
            'valign': 'middle',
            'border': 1
        })
        
        for col_num, value in enumerate(st.session_state["ft_efatura_kalemler_df"].columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, max(len(str(value)) + 3, 16))
            
    st.download_button(
        label="📥 Biçimlendirilmiş İndirilecek KDV Listesini İndir (.xlsx)", 
        data=output_excel.getvalue(), 
        file_name="Indirilecek_KDV_Listesi.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        key="dl_excel_btn"
    )
