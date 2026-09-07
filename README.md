# Alfatih SmartSchool

A standalone desktop application for managing **daily student attendance** (absensi) and **grades** (nilai) at Alfatih SmartSchool. Built with **PySide6 (Qt for Python)** and **SQLite**, fully offline.

![License](https://img.shields.io/badge/license-MIT-green)

---

## Screenshot application
!(Alfatih-SmartSchool)[ss.png]

## Features

### Attendance (Absensi)
- Mark daily attendance per class with statuses: **Hadir** (Present), **Alpa** (Absent), **Izin** (Excused), **Sakit** (Sick).
- "Mark All Present" shortcut for a class on a given date.
- Daily summary chips showing the current tally of each status.
- Record (Rekap) table with filters (class, status, date range, name search).
- Per-student summary (Kesimpulan) with total counts per status.

### Grades (Nilai)
- Grade entry per class, exam type (**Latihan**, **Ulangan**, **UTS**, **UAS**) and date.
- Weighted final grades automatically computed from the configured weights.
- Full grade record table with filters and search.
- Per-student grade summary with weighted average and letter grade.

### Scheduling (Jadwal)
- Schedule exams/exercises per class (type, date, optional note).

### Export
- Export attendance and grade records (plus summaries) to **Excel (.xlsx)** or **CSV**.

### Calendar
- Built-in Indonesian national holiday calendar (2025–2029); holiday dates are highlighted and warned before marking attendance.

---

## Screenshots

_(Add screenshots of the main pages here.)_

---

## Requirements

- Python **3.9+** (tested on 3.11)
- See [`requirements.txt`](requirements.txt):
  - `PySide6 >= 6.6`
  - `openpyxl >= 3.1`

---

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/YogaRmdn/Alfatih-SmartSchool.git
cd Alfatih-SmartSchool

# 2. (Recommended) Create and activate a virtual environment
python -m venv .venv
.venv\Scripts\activate        # Windows
source .venv/bin/activate     # macOS / Linux

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the application
python aplikasi_absensi_qt.py
```

---

## Usage

1. Go to **Kelola Data** (Manage Data) to create classes and add students.
2. On the **Absensi** page, select a class and date, then mark each student's attendance.
3. On the **Input Nilai** page, select class, exam type, and date, fill in each student's grade, and click **Simpan Nilai** (Save Grades).
4. Use **Jadwal** to plan upcoming exams.
5. Review records and summaries on the **Rekap** pages and export to Excel/CSV.

The application stores everything locally in `absensi.db`, which is created automatically next to the script on first run.

---

## Project Structure

```
Alfatih SmartSchool/
├── aplikasi_absensi_qt.py   # Entire application (single file)
├── requirements.txt         # Python dependencies
├── fonts/                   # Poppins font family (bundled)
├── img/
│   ├── logo.png             # App/window icon
│   └── icons/               # Sidebar navigation icons (SVG)
└── absensi.db               # SQLite database (auto-generated)
```

---

## Database

SQLite with five tables, all linked through foreign keys with `ON DELETE CASCADE`:

| Table            | Purpose                              |
|------------------|--------------------------------------|
| `kelas`          | Classes                              |
| `siswa`          | Students (belong to a class)         |
| `presensi`       | Daily attendance records             |
| `jadwal_ujian`   | Exam/exercise schedules              |
| `nilai`          | Grade records per student/type/date  |

All queries use parameterized statements (no SQL injection risk).

---

## Grade Weights

The weighted final grade is computed using the built-in weights map (`BOBOT_NILAI`), which you can adjust at the top of `aplikasi_absensi_qt.py`:

| Type      | Default weight |
|-----------|----------------|
| Latihan   | 20%            |
| Ulangan   | 30%            |
| UTS       | 20%            |
| UAS       | 30%            |

---

## Troubleshooting

- **`ModuleNotFoundError: PySide6`** — dependencies are not installed; run `pip install -r requirements.txt`.
- **Database not found** — the app creates `absensi.db` automatically next to the script on first run; do not move the file while the app is running.
- **Holidays after 2029** — the built-in Indonesian holiday calendar is static; update `LIBUR_TETAP`/`LIBUR_DINAMIS` in the source to extend coverage.

---

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details. *(Add a LICENSE file if you choose to publish under MIT.)*
