import os
import tkinter as tk
from tkinter import messagebox, filedialog
from openpyxl import Workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def get_next_run_number(save_dir):
    existing = [
        d for d in os.listdir(save_dir)
        if os.path.isdir(os.path.join(save_dir, d)) and d.startswith("batch run ")
    ]
    return len(existing) + 1


def on_save_as():
    folder = filedialog.askdirectory(title="Select save location")
    if folder:
        save_dir.set(folder)


def on_run():
    try:
        iterations = int(entry.get())
        if iterations < 1:
            raise ValueError
    except ValueError:
        messagebox.showerror("Invalid input", "Please enter a positive whole number.")
        return

    save_location = save_dir.get()
    if not save_location:
        messagebox.showerror("No save location", "Please select a save location first.")
        return

    run_number = get_next_run_number(save_location)
    folder = os.path.join(save_location, f"batch run {run_number}")
    os.makedirs(folder)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Batch Run {run_number}"
    ws.append(["Iteration", "Output"])

    log_text.delete("1.0", tk.END)

    for i in range(1, iterations + 1):
        message = f"G'day Australia! {i}"
        ws.append([i, message])
        log_text.insert(tk.END, message + "\n")
        log_text.see(tk.END)
        log_text.update()

    filename = f"batch run {run_number}.xlsx"
    wb.save(os.path.join(folder, filename))
    log_text.insert(tk.END, f"\nOutput saved to: {folder}\\{filename}\n")


app = tk.Tk()
app.title("G'day Australia")
app.resizable(False, False)

frame = tk.Frame(app, padx=10, pady=10)
frame.pack()

tk.Label(frame, text="Number of iterations:").grid(row=0, column=0, sticky="w")
entry = tk.Entry(frame, width=10)
entry.insert(0, "100")
entry.grid(row=0, column=1, padx=5)

tk.Button(frame, text="Run", command=on_run, width=10).grid(row=0, column=2)

tk.Label(frame, text="Save location:").grid(row=1, column=0, sticky="w", pady=(10, 0))
save_dir = tk.StringVar(value=BASE_DIR)
tk.Entry(frame, textvariable=save_dir, width=35, state="readonly").grid(row=1, column=1, pady=(10, 0))
tk.Button(frame, text="Save As...", command=on_save_as, width=10).grid(row=1, column=2, pady=(10, 0))

log_text = tk.Text(frame, width=50, height=20, state="normal")
log_text.grid(row=2, column=0, columnspan=3, pady=(10, 0))

app.mainloop()
