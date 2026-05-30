# 🧠 Brainiac Budget Tracker Pro

A beautiful, high-performance desktop application for budget tracking, expense logging, and real-time visual analytics. Built with Python, Tkinter, and Matplotlib.

## ✨ Features
- **Modern Dark UI**: A premium dark-mode interface inspired by the Catppuccin Mocha palette.
- **Real-Time Analytics**: Built-in interactive pie chart displaying categorical expense distribution in real-time.
- **Smart Budget Guard**: Set a monthly budget threshold with live visual status alerts (**Stable** vs **Over Budget**).
- **Persistent Storage**: Robust SQLite backend ensuring all logs and custom categories are saved locally.
- **Premium Excel Ledger Export**: Generates beautifully styled native Excel (.xlsx) sheets with zebra striping, currency formats, and automatic column scaling.
- **Custom Category Strategy**: Add custom expense categories instantly.

## 🚀 Getting Started

### Prerequisites
Make sure you have Python 3 installed on your system. You will also need to install the required external libraries:

```bash
pip install matplotlib openpyxl
```

*Note: Tkinter and SQLite come pre-packaged with standard Python installations.*

### Running the Application
Run the Python script directly from your terminal:

```bash
python budget_tracker.py
```

## 📊 Technologies Used
- **Language**: Python 3
- **GUI Engine**: Tkinter (scaled for high-DPI displays)
- **Analytics & Plots**: Matplotlib (embedded `TkAgg` backend)
- **Database**: SQLite3
- **Export System**: Openpyxl with custom typography and alignments

---
Developed with 🧠 by **BrainiacTech**
