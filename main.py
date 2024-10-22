import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog, messagebox, Tk, Button, Label, Frame, Toplevel, ttk

def select_file():
    file_path = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
        return None
    return file_path

def main():
    root = Tk()
    root.title("Análisis de Publicidad y Ventas")
    root.geometry("400x200")

    frame = Frame(root)
    frame.pack(pady=20)

    label = Label(frame, text="Seleccione un archivo Excel para analizar:", font=("Arial", 12))
    label.pack(pady=10)

    def on_button_click():
        file_path = select_file()
        if file_path:
            analyze_file(file_path)

    button = Button(frame, text="Seleccionar archivo", command=on_button_click, font=("Arial", 12))
    button.pack(pady=10)

    root.mainloop()

def analyze_file(file_path):
    loading_window = Toplevel()
    loading_window.title("Cargando")
    loading_window.geometry("300x100")
    Label(loading_window, text="Cargando archivo, por favor espere...").pack(pady=20)
    progress_bar = ttk.Progressbar(loading_window, mode='indeterminate')
    progress_bar.pack(pady=10)
    progress_bar.start()

    try:
        df = pd.read_excel(file_path)
        loading_window.destroy()
        messagebox.showinfo("Éxito", "Archivo cargado exitosamente.")
    except Exception as e:
        loading_window.destroy()
        messagebox.showerror("Error", f"Error al cargar el archivo: {e}")
        return

    required_columns = ["x", "y"]
    for col in required_columns:
        if col not in df.columns:
            messagebox.showerror("Error", f"La columna '{col}' no está presente en el archivo Excel.")
            return

    x = df['x']
    y = df['y']

    n = len(x)
    Sxx = np.sum(x ** 2) - (np.sum(x) ** 2) / n
    Syy = np.sum(y ** 2) - (np.sum(y) ** 2) / n
    Sxy = np.sum(x * y) - (np.sum(x) * np.sum(y)) / n

    r = Sxy / np.sqrt(Sxx * Syy)
    R_squared = (r ** 2)

    messagebox.showinfo("Resultados", f"Coeficiente de correlación (r): {r}\nCoeficiente de determinación (R^2): {R_squared}")

    plt.scatter(x, y, color='blue', label='Datos')
    plt.xlabel('Publicidad (minutos)', fontsize=12)
    plt.ylabel('Ventas (unidades)', fontsize=12)
    plt.title('Relación entre Publicidad y Ventas', fontsize=14, fontweight='bold')
    plt.grid(True, linestyle='--', alpha=0.7)

    for i in range(len(x)):
        plt.text(x[i], y[i], f'({x[i]}, {y[i]})', fontsize=8, ha='right')

    m, b = np.polyfit(x, y, 1)
    plt.plot(x, m * x + b, color='red', linestyle='-', linewidth=2, label='Línea de mejor ajuste')

    plt.legend()

    def save_plot():
        save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png"), ("All files", "*.*")])
        if save_path:
            plt.savefig(save_path)
            messagebox.showinfo("Guardado", f"Gráfica guardada en {save_path}")

    save_button = Button(plt.gcf().canvas.get_tk_widget(), text="Guardar Gráfica", command=save_plot)
    save_button.pack(side='bottom')

    plt.show()

if __name__ == "__main__":
    main()