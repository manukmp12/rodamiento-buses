import os
import subprocess
import platform
import tkinter as tk
from tkinter import messagebox
import pandas as pd
from datetime import datetime, timedelta

EXCEL_FILE = "rodamiento_modificado.xlsx"
TIEMPO_TURNO = 15 * 60  # 15 minutos en segundos

class RodamientoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Control de Rodamiento")
        self.root.attributes("-topmost", True)
        self.root.geometry("420x330")
        self.root.minsize(360, 290)
        self.root.config(bg="#121212")

        self.df = pd.DataFrame()
        self.indice_actual = 0
        self.tiempo_restante = TIEMPO_TURNO
        self.timer_id = None
        self.en_pausa = True

        # --- BARRA SUPERIOR COMPACTA ---
        frame_top = tk.Frame(root, bg="#121212")
        frame_top.pack(fill=tk.X, padx=8, pady=(4, 2))

        tk.Label(
            frame_top, text="🚚 CONTROL DE RODAMIENTO", 
            font=("Arial", 10, "bold"), bg="#121212", fg="#00ffcc"
        ).pack(side=tk.LEFT)

        # Botón Modificar: abre directamente tu archivo Excel sin tocar fórmulas ni formatos
        tk.Button(
            frame_top, text="✏️ Modificar", command=self.abrir_excel,
            bg="#e67e22", fg="white", font=("Arial", 8, "bold"), relief="flat", padx=6, pady=1
        ).pack(side=tk.RIGHT)

        # ==============================================================
        # --- TARJETA ÚNICA ---
        # ==============================================================
        card = tk.Frame(root, bg="#202020", bd=1, relief="solid")
        card.pack(fill=tk.BOTH, expand=True, padx=8, pady=2)

        # 1. ARRIBA: Despachado (-15 min) desvanecido
        self.lbl_prev = tk.Label(
            card, text="▲ DESPACHADO: ---", 
            font=("Arial", 8), bg="#202020", fg="#666666"
        )
        self.lbl_prev.pack(fill=tk.X, padx=4, pady=(4, 1))

        # Divisor superior sutil
        tk.Frame(card, bg="#2e2e2e", height=1).pack(fill=tk.X, padx=8, pady=1)

        # 2. CENTRO: Vehículo actual
        self.lbl_trayecto = tk.Label(
            card, text="---", 
            font=("Arial", 12, "bold"), bg="#202020", fg="#ffcc00"
        )
        self.lbl_trayecto.pack(pady=0)

        self.lbl_hora_actual = tk.Label(
            card, text="⏰ Salida: ---", 
            font=("Arial", 9, "bold"), bg="#202020", fg="#e0e0e0"
        )
        self.lbl_hora_actual.pack(pady=0)

        self.lbl_vehiculo = tk.Label(
            card, text="---", 
            font=("Arial", 32, "bold"), bg="#202020", fg="#00ffcc"
        )
        self.lbl_vehiculo.pack(pady=0)

        self.lbl_conductor = tk.Label(
            card, text="---", 
            font=("Arial", 11, "bold"), bg="#202020", fg="#ffffff"
        )
        self.lbl_conductor.pack(pady=(0, 2))

        # Divisor inferior sutil
        tk.Frame(card, bg="#2e2e2e", height=1).pack(fill=tk.X, padx=8, pady=1)

        # 3. ABAJO: Próximo turno (+15 min) desvanecido
        self.lbl_next = tk.Label(
            card, text="▼ SIGUIENTE: ---", 
            font=("Arial", 8), bg="#202020", fg="#666666"
        )
        self.lbl_next.pack(fill=tk.X, padx=4, pady=(1, 4))

        # ==============================================================
        # --- BARRA INFERIOR (Timer + Botones) ---
        # ==============================================================
        bar_controls = tk.Frame(root, bg="#121212")
        bar_controls.pack(fill=tk.X, padx=8, pady=(2, 6))

        # Cronómetro a la izquierda
        self.lbl_timer = tk.Label(
            bar_controls, text="15:00", 
            font=("Arial", 14, "bold"), bg="#121212", fg="#00ff00", width=5
        )
        self.lbl_timer.pack(side=tk.LEFT)

        # Botones a la derecha
        self.btn_siguiente = tk.Button(
            bar_controls, text="⏭ Siguiente", command=self.siguiente_vehiculo, 
            bg="#17a2b8", fg="white", font=("Arial", 8, "bold"), relief="flat", padx=6, pady=2
        )
        self.btn_siguiente.pack(side=tk.RIGHT, padx=2)

        self.btn_pausa = tk.Button(
            bar_controls, text="⏸ Pausa", command=self.pausar_cronometro, 
            bg="#ffc107", fg="black", font=("Arial", 8, "bold"), relief="flat", padx=6, pady=2
        )
        self.btn_pausa.pack(side=tk.RIGHT, padx=2)

        self.btn_iniciar = tk.Button(
            bar_controls, text="▶ Iniciar", command=self.iniciar_cronometro, 
            bg="#28a745", fg="white", font=("Arial", 8, "bold"), relief="flat", padx=6, pady=2
        )
        self.btn_iniciar.pack(side=tk.RIGHT, padx=2)

        self.cargar_datos()

    def abrir_excel(self):
        """Abre directamente la hoja de cálculo en el programa nativo de Excel."""
        if not os.path.exists(EXCEL_FILE):
            messagebox.showerror("Error", f"No se encontró el archivo:\n{EXCEL_FILE}")
            return

        try:
            sistema = platform.system()
            if sistema == "Windows":
                os.startfile(EXCEL_FILE)
            elif sistema == "Darwin":  # macOS
                subprocess.call(["open", EXCEL_FILE])
            else:  # Linux
                subprocess.call(["xdg-open", EXCEL_FILE])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
            return

        # Aviso para sincronizar una vez editado en Excel
        recargar = messagebox.askyesno(
            "Excel abierto",
            "Modifica los datos que necesites directamente en Excel.\n\n"
            "Guarda los cambios (Ctrl + G) y ciérralo.\n"
            "¿Deseas recargar la aplicación ahora para aplicar los cambios?"
        )
        if recargar:
            self.cargar_datos()

    def cargar_datos(self):
        """Lee el Excel en modo solo lectura (conservando intactas fórmulas y estilos)."""
        if not os.path.exists(EXCEL_FILE):
            messagebox.showwarning("Atención", f"No se encontró {EXCEL_FILE}")
            return
        try:
            self.df = pd.read_excel(EXCEL_FILE)
            if self.df.empty:
                messagebox.showwarning("Vacío", "El archivo no contiene registros.")
                return
            if self.indice_actual >= len(self.df):
                self.indice_actual = 0
            self.mostrar_vehiculo_actual()
        except PermissionError:
            messagebox.showerror("Error", "Cierra el archivo en Excel antes de recargar para que no esté bloqueado.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")

    def obtener_numero_interno(self, fila):
        if "N° INTERNO" in fila and pd.notna(fila["N° INTERNO"]):
            val = fila["N° INTERNO"]
            return int(val) if isinstance(val, (int, float)) and val == int(val) else val
        if "NUMERO INTERNO" in fila and pd.notna(fila["NUMERO INTERNO"]):
            val = fila["NUMERO INTERNO"]
            return int(val) if isinstance(val, (int, float)) and val == int(val) else val
        return "N/D"

    def parsear_hora(self, valor_hora):
        try:
            str_h = str(valor_hora).strip()
            if len(str_h) > 8 and " " in str_h:
                str_h = str_h.split()[1]
            return datetime.strptime(str_h[:8], "%H:%M:%S")
        except:
            return None

    def formatear_hora(self, valor_hora):
        t = self.parsear_hora(valor_hora)
        return t.strftime("%I:%M %p") if t else str(valor_hora)

    def calcular_hora_desplazada(self, hora_base_raw, delta_minutos):
        t = self.parsear_hora(hora_base_raw)
        if t:
            return (t + timedelta(minutes=delta_minutos)).strftime("%I:%M %p")
        return "---"

    def mostrar_vehiculo_actual(self):
        total = len(self.df)
        if self.df.empty or self.indice_actual >= total:
            self.lbl_trayecto.config(text="---")
            self.lbl_hora_actual.config(text="⏰ Salida: ---")
            self.lbl_vehiculo.config(text="---")
            self.lbl_conductor.config(text="---")
            self.lbl_prev.config(text="▲ DESPACHADO: Sin registros")
            self.lbl_next.config(text="▼ SIGUIENTE: Sin registros")
            return

        # Turno actual
        fila = self.df.iloc[self.indice_actual]
        hora_raw = fila.get("HORA", "N/D")
        self.lbl_trayecto.config(text=str(fila.get("TRAYECTO", "N/D")))
        self.lbl_hora_actual.config(text=f"⏰ Salida: {self.formatear_hora(hora_raw)}")
        self.lbl_vehiculo.config(text=str(self.obtener_numero_interno(fila)))
        self.lbl_conductor.config(text=str(fila.get("NOMBRE DEL CONDUCTOR", "N/D")))

        # Carro anterior (-15m)
        if self.indice_actual > 0:
            f_prev = self.df.iloc[self.indice_actual - 1]
            h_prev = self.formatear_hora(f_prev.get("HORA", ""))
            if not h_prev or h_prev == "N/D":
                h_prev = self.calcular_hora_desplazada(hora_raw, -15)
            self.lbl_prev.config(
                text=f"▲ DESPACHADO ({h_prev}): #{self.obtener_numero_interno(f_prev)} | {f_prev.get('NOMBRE DEL CONDUCTOR', '')}"
            )
        else:
            self.lbl_prev.config(text="▲ DESPACHADO: (Inicio de jornada)")

        # Carro siguiente (+15m)
        if self.indice_actual + 1 < total:
            f_next = self.df.iloc[self.indice_actual + 1]
            h_next = self.formatear_hora(f_next.get("HORA", ""))
            if not h_next or h_next == "N/D":
                h_next = self.calcular_hora_desplazada(hora_raw, 15)
            self.lbl_next.config(
                text=f"▼ SIGUIENTE ({h_next}): #{self.obtener_numero_interno(f_next)} | {f_next.get('NOMBRE DEL CONDUCTOR', '')}"
            )
        else:
            self.lbl_next.config(text="▼ SIGUIENTE: (Fin del rodamiento)")

    def iniciar_cronometro(self):
        if self.df.empty:
            messagebox.showwarning("Advertencia", "No hay datos cargados.")
            return
        if self.en_pausa:
            self.en_pausa = False
            self.actualizar_cronometro()

    def pausar_cronometro(self):
        self.en_pausa = True
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None

    def siguiente_vehiculo(self):
        if self.df.empty:
            return
        self.pausar_cronometro()
        self.indice_actual += 1
        if self.indice_actual < len(self.df):
            self.tiempo_restante = TIEMPO_TURNO
            self.mostrar_vehiculo_actual()
            self.lbl_timer.config(text="15:00", fg="#00ff00")
            self.iniciar_cronometro()
        else:
            self.mostrar_vehiculo_actual()
            self.lbl_timer.config(text="00:00", fg="yellow")
            messagebox.showinfo("Aviso", "Se han completado todos los despachos.")

    def actualizar_cronometro(self):
        if self.en_pausa:
            return
        if self.tiempo_restante > 0:
            m = self.tiempo_restante // 60
            s = self.tiempo_restante % 60
            self.lbl_timer.config(text=f"{m:02d}:{s:02d}")
            self.lbl_timer.config(fg="#ff4444" if self.tiempo_restante <= 120 else "#00ff00")
            self.tiempo_restante -= 1
            self.timer_id = self.root.after(1000, self.actualizar_cronometro)
        else:
            self.siguiente_vehiculo()

if __name__ == "__main__":
    root = tk.Tk()
    app = RodamientoApp(root)
    root.mainloop()