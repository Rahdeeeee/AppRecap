import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import streamlit as st

# Konfigurasi halaman web
st.set_page_config(
    page_title="Dashboard Absensi Berbasis ID", page_icon="📊", layout="wide"
)

st.title("📊 DASHBOARD & REKAP ABSENSI (BERBASIS ID)")
st.write(
    "Unggah file Excel absensi mentah. Sistem akan mencocokkan data"
    " berdasarkan ID Karyawan secara otomatis."
)

# 1. Upload File Excel Mentah
uploaded_file = st.file_uploader(
    "Pilih file Excel Absensi (.xlsx / .xls)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
  try:
    # Membaca data Excel
    df = pd.read_excel(uploaded_file)

    # Memastikan kolom penting tersedia (termasuk ID)
    if "Waktu Absensi" in df.columns and "ID" in df.columns:
      # Pisahkan Waktu Absensi menjadi Tanggal dan Jam
      df["Waktu Absensi"] = pd.to_datetime(df["Waktu Absensi"])
      df["Tanggal"] = df["Waktu Absensi"].dt.strftime("%d/%m/%Y")
      df["Jam"] = df["Waktu Absensi"].dt.strftime("%H:%M:%S")

      # Urutkan berdasarkan ID dan Waktu
      df = df.sort_values(by=["ID", "Waktu Absensi"])

      # Proses Detail Per Hari (untuk Sheet 1)
      rekap_detail = []
      # Mengelompokkan berdasarkan ID, Nama, Tanggal, Departemen, Posisi
      grouped = df.groupby(["ID", "Nama", "Tanggal", "Departemen", "Posisi"])

      for (id_karyawan, nama, tanggal, dept, posisi), group in grouped:
        jam_list = sorted(group["Jam"].tolist())

        # Deteksi Jam Masuk dan Keluar
        jam_masuk = ""
        jam_keluar = ""

        if len(jam_list) == 1:
          waktu_str = jam_list[0]
          jam_obj = pd.to_datetime(waktu_str, format="%H:%M:%S").time()
          if jam_obj.hour < 12:
            jam_masuk = waktu_str
          else:
            jam_keluar = waktu_str
        elif len(jam_list) >= 2:
          jam_masuk = jam_list[0]
          jam_keluar = jam_list[-1]

        # Tentukan Flag No Posting In & Out
        no_posting_in = 1 if jam_masuk == "" else 0
        no_posting_out = 1 if jam_keluar == "" else 0
        hadir_flag = 1 if (jam_masuk != "" or jam_keluar != "") else 0

        rekap_detail.append({
            "ID": id_karyawan,
            "Nama": nama,
            "Tanggal Masuk": tanggal,
            "Jam Masuk": jam_masuk,
            "Tanggal Keluar": tanggal,
            "Jam Keluar": jam_keluar,
            "Departemen": dept,
            "Posisi": posisi,
            "No_Posting_In_Flag": no_posting_in,
            "No_Posting_Out_Flag": no_posting_out,
            "Hadir_Flag": hadir_flag,
        })

      df_detail = pd.DataFrame(rekap_detail)

      # Buat Rekap Ringkasan per ID (Nama, Departemen, Posisi mengikuti berdasarkan ID)
      df_summary = (
          df_detail.groupby(["ID", "Nama", "Departemen", "Posisi"])
          .agg(
              Total_Hadir=("Hadir_Flag", "sum"),
              No_Posting_In=("No_Posting_In_Flag", "sum"),
              No_Posting_Out=("No_Posting_Out_Flag", "sum"),
          )
          .reset_index()
      )

      # Urutkan berdasarkan Departemen lalu ID secara A-Z
      df_summary = df_summary.sort_values(
          by=["Departemen", "ID"], ascending=[True, True]
      ).reset_index(drop=True)

      # Persentase Kehadiran
      TARGET_HARI_KERJA = 22  # Standar target hari kerja sebulan
      df_summary["% Kehadiran"] = (
          (df_summary["Total_Hadir"] / TARGET_HARI_KERJA) * 100
      ).round(1)

      # Urutan kolom Dashboard
      df_dashboard = df_summary[[
          "ID",
          "Nama",
          "Departemen",
          "Posisi",
          "Total_Hadir",
          "No_Posting_In",
          "No_Posting_Out",
          "% Kehadiran",
      ]]
      df_dashboard.rename(
          columns={"Total_Hadir": "Total Kehadiran"}, inplace=True
      )
      df_dashboard.insert(0, "No", range(1, len(df_dashboard) + 1))

      # --- TAMPILAN DI WEB STREAMLIT ---
      st.markdown("---")
      total_hadir_all = df_dashboard["Total Kehadiran"].sum()
      total_np_in_all = df_dashboard["No_Posting_In"].sum()
      total_np_out_all = df_dashboard["No_Posting_Out"].sum()

      # Kartu Ringkasan (KPI Cards)
      c1, c2, c3 = st.columns(3)
      c1.metric(label="✅ TOTAL KEHADIRAN", value=total_hadir_all)
      c2.metric(label="⚠️ TOTAL NO POSTING IN", value=total_np_in_all)
      c3.metric(label="⚠️ TOTAL NO POSTING OUT", value=total_np_out_all)

      st.markdown("---")
      st.subheader("📋 TABEL DASHBOARD RINGKASAN (Berdasarkan ID)")
      st.dataframe(df_dashboard, use_container_width=True)

      st.subheader("📄 DETAIL ABSENSI HARIAN")
      st.dataframe(
          df_detail.drop(
              columns=[
                  "No_Posting_In_Flag",
                  "No_Posting_Out_Flag",
                  "Hadir_Flag",
              ]
          ),
          use_container_width=True,
      )

      # 2. Proses Export ke Excel (2 Sheet: Detail & Dashboard)
      output = io.BytesIO()
      wb = Workbook()

      # --- SHEET 1: DETAIL ABSENSI ---
      ws_detail = wb.active
      ws_detail.title = "Detail Absensi"
      ws_detail.views.sheetView[0].showGridLines = True

      detail_headers = [
          "ID",
          "Nama",
          "Tanggal Masuk",
          "Jam Masuk",
          "Tanggal Keluar",
          "Jam Keluar",
          "Departemen",
          "Posisi",
      ]
      ws_detail.append(detail_headers)
      for col_num in range(1, len(detail_headers) + 1):
        cell = ws_detail.cell(row=1, column=col_num)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            start_color="1F497D", end_color="1F497D", fill_type="solid"
        )

      for _, row in df_detail.iterrows():
        ws_detail.append([
            row["ID"],
            row["Nama"],
            row["Tanggal Masuk"],
            row["Jam Masuk"],
            row["Tanggal Keluar"],
            row["Jam Keluar"],
            row["Departemen"],
            row["Posisi"],
        ])

      # --- SHEET 2: DASHBOARD & REKAP ---
      ws_dash = wb.create_sheet(title="Dashboard")
      ws_dash.views.sheetView[0].showGridLines = True

      # Judul Dashboard di Excel
      ws_dash.merge_cells("A1:H1")
      ws_dash["A1"] = (
          "DASHBOARD REKAP ABSENSI BERBASIS ID (NO POSTING IN & OUT)"
      )
      ws_dash["A1"].font = Font(
          name="Arial", size=12, bold=True, color="FFFFFF"
      )
      ws_dash["A1"].fill = PatternFill(
          start_color="1F497D", end_color="1F497D", fill_type="solid"
      )
      ws_dash["A1"].alignment = Alignment(horizontal="center", vertical="center")

      # Kotak Ringkasan Angka di Bagian Atas Excel (Baris 3)
      ws_dash["A3"] = "TOTAL KEHADIRAN:"
      ws_dash["B3"] = total_hadir_all
      ws_dash["D3"] = "TOTAL NO POSTING IN:"
      ws_dash["E3"] = total_np_in_all

      ws_dash["A4"] = "TOTAL NO POSTING OUT:"
      ws_dash["B4"] = total_np_out_all

      # Header Tabel Dashboard (Baris 6)
      dash_headers = [
          "No",
          "ID Karyawan",
          "Nama Karyawan",
          "Departemen",
          "Posisi",
          "Total Kehadiran",
          "No Posting In",
          "No Posting Out",
          "% Kehadiran",
      ]

      ws_dash.append([])  # Baris 2 kosong
      ws_dash.append([])  # Baris 3 kosong (Ringkasan)
      ws_dash.append([])  # Baris 4 kosong (Ringkasan)
      ws_dash.append([])  # Baris 5 kosong
      ws_dash.append(dash_headers)  # Baris 6 Header

      # Styling Header Tabel Dashboard
      for col_num in range(1, len(dash_headers) + 1):
        cell = ws_dash.cell(row=6, column=col_num)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            start_color="366092", end_color="366092", fill_type="solid"
        )
        cell.alignment = Alignment(horizontal="center", vertical="center")

      # Masukkan Data Ringkasan Dashboard ke Excel
      for _, row in df_dashboard.iterrows():
        ws_dash.append([
            row["No"],
            row["ID"],
            row["Nama"],
            row["Departemen"],
            row["Posisi"],
            row["Total Kehadiran"],
            row["No_Posting_In"],
            row["No_Posting_Out"],
            f"{row['% Kehadiran']}%",
        ])

      wb.save(output)
      processed_data = output.getvalue()

      # 3. Tombol Download Excel
      st.download_button(
          label=(
              "📥 Download Excel Berbasis ID (Sheet 1: Detail & Sheet 2:"
              " Dashboard)"
          ),
          data=processed_data,
          file_name="Dashboard_Absensi_Berbasis_ID.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

    else:
      st.error(
          "Format kolom pada Excel tidak sesuai! Pastikan terdapat kolom 'ID' dan"
          " 'Waktu Absensi'."
      )

  except Exception as e:
    st.error(f"Terjadi kesalahan saat memproses file: {e}")
