import os
import pandas as pd
import plotly.express as px
import streamlit as st

# Sayfa Yapılandırması
st.set_page_config(page_title="Lead Core Bulut Yönetim Paneli", layout="wide")

# --- 🎯 METRİK MOTORUNU BOZMAYAN KESİN HAREKETLİ ARKA PLAN (HTML BILESENI) ---
# Süslü parantez uyuşmazlığını önlemek için saf HTML tüneli kullanılmıştır
st.components.v1.html("""
<style>
body, html { margin: 0; padding: 0; }
.bg-gif {
    position: fixed;
    top: 0; left: 0; width: 100vw; height: 100vh;
    background-image: url("https://giphy.com");
    background-size: cover;
    background-position: center;
    background-repeat: no-repeat;
    z-index: -2;
}
</style>
<div class="bg-gif"></div>
""", height=0)

# Streamlit arayüz bileşenlerini şeffaflaştıran ve ana paneli okutan güvenli CSS (Süslü parantezsiz, saf metin formatlama)
st.markdown("<style>.stApp { background: transparent; } .main .block-container { background-color: rgba(255, 255, 255, 0.94); padding: 40px; border-radius: 16px; box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1); margin-top: 20px; }</style>", unsafe_html=True)

# --- 🖼️ LOGO VE BAŞLIK ALANI ---
if os.path.exists("fabrika_logo.png"):
    st.image("fabrika_logo.png", width=250)

st.title("Lead Core Bulut Performans Analiz Paneli")
st.markdown("GitHub deponuz üzerinden doğrudan okunan, telefondan erişilebilir kesintisiz bulut ekranı.")

# --- 🎯 DOĞRUDAN DEPO İÇİ OKUMA AYARI ---
tam_yol = "Lead_Core_Pres_Üretim.xlsx"

@st.cache_data(ttl=600)
def yeni_excel_mimarisi_oku():
    try:
        excel_dosyasi = pd.ExcelFile(tam_yol, engine="openpyxl")
        tum_satirlar = []
        
        toplam_aktif_gun_sayisi = len(excel_dosyasi.sheet_names)
        maks_teorik_vardiya = toplam_aktif_gun_sayisi * 7 * 3
        vardiya_hedef_adet = 90000.0

        for sayfa in excel_dosyasi.sheet_names:
            df_sayfa = pd.read_excel(tam_yol, sheet_name=sayfa, header=None, engine="openpyxl")
            
            istisnalar = {"Sabah": {}, "Akşam": {}, "Gece": {}}
            for r in range(49, min(56, len(df_sayfa))):
                m_s, k_s = df_sayfa.iloc[r, 9], df_sayfa.iloc[r, 10]
                if pd.notna(m_s) and pd.notna(k_s):
                    try: istisnalar["Sabah"][int(float(m_s))] = str(k_s).strip().upper()
                    except: pass
                    
                m_a, k_a = df_sayfa.iloc[r, 11], df_sayfa.iloc[r, 13]
                if pd.notna(m_a) and pd.notna(k_a):
                    try: istisnalar["Akşam"][int(float(m_a))] = str(k_a).strip().upper()
                    except: pass
                    
                m_g, k_g = df_sayfa.iloc[r, 14], df_sayfa.iloc[r, 15]
                if pd.notna(m_g) and pd.notna(k_g):
                    try: istisnalar["Gece"][int(float(m_g))] = str(k_g).strip().upper()
                    except: pass

            for m_idx in range(7):
                makine_no = m_idx + 1
                bas_satir = 5 + (m_idx * 6)
                
                yasakli = ["-", "YOK", "BOŞ", "BOS", "OP-YOK", "OP YOK", "NAN", "NONE", ""]
                
                ops_sabah = [df_sayfa.iloc[bas_satir, 1], df_sayfa.iloc[bas_satir+1, 1]]
                ops_aksam = [df_sayfa.iloc[bas_satir+2, 1], df_sayfa.iloc[bas_satir+3, 1]]
                ops_gece  = [df_sayfa.iloc[bas_satir+4, 1], df_sayfa.iloc[bas_satir+5, 1]]
                
                ops_sabah = [str(o).strip() for o in ops_sabah if pd.notna(o) and str(o).strip() != "" and str(o).strip().upper() not in yasakli]
                ops_aksam = [str(o).strip() for o in ops_aksam if pd.notna(o) and str(o).strip() != "" and str(o).strip().upper() not in yasakli]
                ops_gece  = [str(o).strip() for o in ops_gece if pd.notna(o) and str(o).strip() != "" and str(o).strip().upper() not in yasakli]
                
                kod_s = istisnalar["Sabah"].get(makine_no, "NORMAL")
                kod_a = istisnalar["Akşam"].get(makine_no, "NORMAL")
                kod_g = istisnalar["Gece"].get(makine_no, "NORMAL")
                
                h_sabah_sabit = 0.0 if kod_s in ["T-ARIZA", "M-YOK", "OP-YOK"] else vardiya_hedef_adet
                h_aksam_sabit = 0.0 if kod_a in ["T-ARIZA", "M-YOK", "OP-YOK"] else vardiya_hedef_adet
                h_gece_sabit  = 0.0 if kod_g in ["T-ARIZA", "M-YOK", "OP-YOK"] else vardiya_hedef_adet
                
                toplam_urt_sabah, toplam_urt_aksam, toplam_urt_gece = 0.0, 0.0, 0.0
                for p_idx in range(0, 5, 2):
                    r_aktif = bas_satir + p_idx
                    v_s = pd.to_numeric(df_sayfa.iloc[r_aktif, 3], errors='coerce')
                    v_a = pd.to_numeric(df_sayfa.iloc[r_aktif, 4], errors='coerce')
                    v_g = pd.to_numeric(df_sayfa.iloc[r_aktif, 5], errors='coerce')
                    toplam_urt_sabah += 0.0 if pd.isna(v_s) else float(v_s)
                    toplam_urt_aksam += 0.0 if pd.isna(v_a) else float(v_a)
                    toplam_urt_gece  += 0.0 if pd.isna(v_g) else float(v_g)
                
                for p_idx in range(0, 5, 2):
                    r_aktif = bas_satir + p_idx
                    urun_adi = df_sayfa.iloc[r_aktif, 2]
                    urun_adi = "Belirtilmemiş Ürün" if pd.isna(urun_adi) or str(urun_adi).strip() == "" else str(urun_adi).strip()
                    
                    urt_s = pd.to_numeric(df_sayfa.iloc[r_aktif, 3], errors='coerce')
                    urt_a = pd.to_numeric(df_sayfa.iloc[r_aktif, 4], errors='coerce')
                    urt_g = pd.to_numeric(df_sayfa.iloc[r_aktif, 5], errors='coerce')
                    
                    urt_s = 0.0 if pd.isna(urt_s) else float(urt_s)
                    urt_a = 0.0 if pd.isna(urt_a) else float(urt_a)
                    urt_g = 0.0 if pd.isna(urt_g) else float(urt_g)
                    
                    if ops_sabah and (urt_s > 0 or (kod_s == "OP-YOK" and p_idx == 0)):
                        op_sayisi = len(ops_sabah)
                        urt_bolunmus = urt_s / op_sayisi
                        h_sabah = 0.0 if kod_s in ["T-ARIZA", "M-YOK", "OP-YOK"] else (vardiya_hedef_adet if p_idx == 0 else 0.0)
                        for op in ops_sabah:
                            tum_satirlar.append({
                                "Tarih": str(sayfa), "Makine_No": f"{makine_no}. Machine", "Vardiya": "Sabah",
                                "Operator_Adi": op, "Urun_Cesidi": urun_adi if urt_s > 0 else "Personel Eksik (Kapalı)",
                                "Uretim": urt_bolunmus, "Ham_Uretim": urt_s, "Durum_Kodu": kod_s, 
                                "Vardiya_Durum_Hedefi": h_sabah_sabit, "Toplam_Aktif_Gun": toplam_aktif_gun_sayisi, "Maks_Vardiya_Kapasite": maks_teorik_vardiya
                            })
                            
                    if ops_aksam and (urt_a > 0 or (kod_a == "OP-YOK" and p_idx == 0)):
                        op_sayisi = len(ops_aksam)
                        urt_bolunmus = urt_a / op_sayisi
                        h_aksam = 0.0 if kod_a in ["T-ARIZA", "M-YOK", "OP-YOK"] else (vardiya_hedef_adet if p_idx == 0 else 0.0)
                        for op in ops_aksam:
                            tum_satirlar.append({
                                "Tarih": str(sayfa), "Makine_No": f"{makine_no}. Machine", "Vardiya": "Akşam",
                                "Operator_Adi": op, "Urun_Cesidi": urun_adi if urt_a > 0 else "Personel Eksik (Kapalı)",
                                "Uretim": urt_bolunmus, "Ham_Uretim": urt_a, "Durum_Kodu": kod_a, 
                                "Vardiya_Durum_Hedefi": h_aksam_sabit, "Toplam_Aktif_Gun": toplam_aktif_gun_sayisi, "Maks_Vardiya_Kapasite": maks_teorik_vardiya
                            })
                            
                    if ops_gece and (urt_g > 0 or (kod_g == "OP-YOK" and p_idx == 0)):
                        op_sayisi = len(ops_gece)
                        urt_bolunmus = urt_g / op_sayisi
                        h_gece = 0.0 if kod_g in ["T-ARIZA", "M-YOK", "OP-YOK"] else (vardiya_hedef_adet if p_idx == 0 else 0.0)
                        for op in ops_gece:
                            tum_satirlar.append({
                                "Tarih": str(sayfa), "Makine_No": f"{makine_no}. Machine", "Vardiya": "Gece",
                                "Operator_Adi": op, "Urun_Cesidi": urun_adi if urt_g > 0 else "Personel Eksik (Kapalı)",
                                "Uretim": urt_bolunmus, "Ham_Uretim": urt_g, "Durum_Kodu": kod_g, 
                                "Vardiya_Durum_Hedefi": h_gece_sabit, "Toplam_Aktif_Gun": toplam_aktif_gun_sayisi, "Maks_Vardiya_Kapasite": maks_teorik_vardiya
                            })

        if not tum_satirlar: return None
        return pd.DataFrame(tum_satirlar)
    except Exception as e:
        st.error(f"Excel dosyası deponun içinden okunurken hata oluştu: {e}")
        return None

df = yeni_excel_mimarisi_oku()
if df is None:
    st.error("❌ Bulut Dosyası Bulunamadı! Lütfen Excel dosyasının GitHub deponuza doğru yüklendiğinden emin olun.")
else:
    # --- YAN PANEL ---
    st.sidebar.header("🔍 Bulut Yönetim Paneli")
    operatorler = sorted(df["Operator_Adi"].unique().tolist())
    secilen_operator = st.sidebar.selectbox("Detaylı Karnesini İncelemek İçin Operatör Seçin:", operatorler)
    
    if st.sidebar.button("🔄 Verileri Depodan Şimdi Yenile"):
        st.cache_data.clear()
        st.rerun()
        
    toplam_aktif_gun = int(df["Toplam_Aktif_Gun"].mean())
    maks_vardiya_kapasite = int(df["Maks_Vardiya_Kapasite"].mean())
    
    tekil_ham_kayitlar = df.groupby(["Tarih", "Makine_No", "Vardiya", "Urun_Cesidi"])["Ham_Uretim"].first().reset_index()
    büyük_fabrika_toplam_uretim = int(tekil_ham_kayitlar["Ham_Uretim"].sum())
    
    tekil_is_emri_df = df.groupby(["Operator_Adi", "Tarih", "Makine_No", "Vardiya"]).agg(
        Net_Vardiya_Uretimi=("Uretim", "sum"),
        Net_Vardiya_Hedefi=("Vardiya_Durum_Hedefi", "first")
    ).reset_index()
    
    op_katilim = tekil_is_emri_df.groupby("Operator_Adi").size().reset_index(name="Calisilan_Toplam_Vardiya")
    op_katilim["Ise_Katilim_Orani_Yuzde"] = (op_katilim["Calisilan_Toplam_Vardiya"] / maks_vardiya_kapasite) * 100
    op_katilim["Ise_Katilim_Orani_Yuzde"] = op_katilim["Ise_Katilim_Orani_Yuzde"].round(1)
    
    op_uretimler = df.groupby("Operator_Adi")["Uretim"].sum().reset_index(name="Adil_Net_Uretim_Katkisi")
    karne_df = pd.merge(op_uretimler, op_katilim, on="Operator_Adi")
    karne_df["Adil_Net_Uretim_Katkisi"] = karne_df["Adil_Net_Uretim_Katkisi"].round(0)

    df_op_veri = df[df["Operator_Adi"] == secilen_operator]
    op_row_data = karne_df[karne_df["Operator_Adi"] == secilen_operator]
    gelinen_gun = int(df_op_veri["Tarih"].nunique()) if not df_op_veri.empty else 0
    op_net_pay_toplam = int(df_op_veri["Uretim"].sum()) if not df_op_veri.empty else 0
    
    makineler_listesi = ", ".join(sorted(df_op_veri["Makine_No"].unique().tolist())) if not df_op_veri.empty else "Yok"
    durum_sayilari = df_op_veri[df_op_veri["Durum_Kodu"] != "NORMAL"].groupby("Durum_Kodu").size().to_dict() if not df_op_veri.empty else {}
    durum_ozet_metni = " / ".join([f"{k}: {v} Kez" for k, v in durum_sayilari.items()]) if durum_sayilari else "Yok"

    # --- 👤 SEÇİLEN OPERATÖRÜN DİJİTAL KARNESİ ---
    st.subheader(f"📋 {secilen_operator} Bulut Performans Karnesi")
    
    if not op_row_data.empty:
        karne_dict = op_row_data.to_dict(orient="records")
        
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("🏢 Fabrika Genel Toplam Üretim", f"{büyük_fabrika_toplam_uretim:,} Adet")
        k2.metric("📦 Pay Edilmiş Net Üretim Katkısı", f"{int(karne_dict['Adil_Net_Uretim_Katkisi']):,} Adet")
        k3.metric("📅 Toplam Üretim Gün Havuzu", f"{gelinen_gun} / {toplam_aktif_gun} Gün", f"{int(karne_dict['Calisilan_Toplam_Vardiya'])} Vardiya")
        k4.metric("📊 Toplam Kapasite Katılım Oranı", f"%{karne_dict['Ise_Katilim_Orani_Yuzde']}")
    else:
        st.info("Seçilen operatöre ait bulut karne verisi hesaplanamadı.")
    
    st.info(f"🤖 **Görev Aldığı Makineler:** {makineler_listesi} | **Excel Toplam Aktif Gün Kümesi:** {toplam_aktif_gun} Gün")
    st.warning(f"⚠️ **Performansı Etkileyen Durum Logları:** {durum_ozet_metni}")

    st.markdown("---")

    # --- GENEL GRAFİKLER SEKMESİ ---
    st.subheader("📊 Fabrika Geneli Karşılaştırma Panoları")
    tab1, tab2 = st.tabs(["📊 Katılım & Üretim Katkı Analizleri", "📋 Filtrelenmiş Günlük Ham Veri Logları"])

    with tab1:
        col_g1, col_g2 = st.columns(2)

        with col_g1:
            st.markdown("##### 📅 Operatörlerin Toplam Gün Havuzuna Katılım Oranları (%)")
            fig_katilim = px.bar(
                karne_df.sort_values(by="Ise_Katilim_Orani_Yuzde", ascending=False),
                x="Operator_Adi", y="Ise_Katilim_Orani_Yuzde", text_auto=True,
                labels={"Operator_Adi": "Operatör", "Ise_Katilim_Orani_Yuzde": "Kapasite Katılımı (%)"},
                template="plotly_white", color="Ise_Katilim_Orani_Yuzde", color_continuous_scale="Purples"
            )
            st.plotly_chart(fig_katilim, use_container_width=True)

        with col_g2:
            st.markdown("##### 📦 Operatörlerin Toplam Üretime Net Adet Katkısı (Bölünmüş Paylar)")
            fig_katki_bar = px.bar(
                karne_df.sort_values(by="Adil_Net_Uretim_Katkisi", ascending=False),
                x="Operator_Adi", y="Adil_Net_Uretim_Katkisi", text_auto=".3s",
                labels={"Operator_Adi": "Operatör", "Adil_Net_Uretim_Katkisi": "Net Üretim Katkısı (Adet)"},
                template="plotly_white", color="Adil_Net_Uretim_Katkisi", color_continuous_scale="RdYlGn"
            )
            st.plotly_chart(fig_katki_bar, use_container_width=True)

        st.markdown("---")
        st.subheader("🤖 Makinelere Göre Toplam Üretim ve Ürün Çeşidi Dağılımı")
        makine_urun_df = df.groupby(["Makine_No", "Urun_Cesidi"])["Ham_Uretim"].sum().reset_index()
        fig_column_urun = px.bar(
            makine_urun_df, x="Makine_No", y="Ham_Uretim", color="Urun_Cesidi",
            labels={"Makine_No": "Pres Makinesi", "Ham_Uretim": "Toplam Üretim (Adet)", "Urun_Cesidi": "Ürün Bilgisi"},
            template="plotly_white", barmode="stack"
        )
        st.plotly_chart(fig_column_urun, use_container_width=True)

    with tab2:
        st.subheader(f"🗂️ {secilen_operator} Günlük Çalışma Günlüğü")
        st.dataframe(
            df_op_veri[["Tarih", "Vardiya", "Makine_No", "Urun_Cesidi", "Uretim", "Ham_Uretim", "Durum_Kodu"]].sort_values(by="Tarih"),
            use_container_width=True
        )


