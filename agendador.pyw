import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import time
import threading
import subprocess
import os
import sys
from datetime import datetime, timedelta

# --- CONFIGURAÇÃO DE AMBIENTE ---
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

ARQUIVO_DB = os.path.join(application_path, "tarefas.json")
ARQUIVO_LOG = os.path.join(application_path, "log_execucao.txt")

class AgendadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Agendador Master 3.3")
        self.root.geometry("1000x600")
        
        self.tarefas = []
        
        # --- 1. ÁREA DE CADASTRO ---
        frame_top = tk.LabelFrame(root, text="Configurar Nova Tarefa", padx=15, pady=15)
        frame_top.pack(fill="x", padx=10, pady=5)
        
        # Linha 1
        tk.Label(frame_top, text="Nome da Tarefa:").grid(row=0, column=0, sticky="w")
        self.entry_nome = tk.Entry(frame_top, width=30)
        self.entry_nome.grid(row=0, column=1, padx=5, sticky="w")
        
        tk.Button(frame_top, text="Selecionar Executável...", command=self.buscar_arquivo).grid(row=0, column=2, padx=5)
        self.entry_path = tk.Entry(frame_top, width=40)
        self.entry_path.grid(row=0, column=3, padx=5)

        ttk.Separator(frame_top, orient='horizontal').grid(row=1, column=0, columnspan=4, sticky="ew", pady=15)

        # Linha 2
        tk.Label(frame_top, text="Começar dia (DD/MM/AAAA):").grid(row=2, column=0, sticky="w")
        self.entry_data = tk.Entry(frame_top, width=15)
        self.entry_data.grid(row=2, column=1, sticky="w", padx=5)
        self.entry_data.insert(0, datetime.now().strftime("%d/%m/%Y"))

        tk.Label(frame_top, text="Hora (HH:MM):").grid(row=2, column=2, sticky="e")
        self.entry_hora = tk.Entry(frame_top, width=10)
        self.entry_hora.grid(row=2, column=3, sticky="w", padx=5)
        self.entry_hora.insert(0, datetime.now().strftime("%H:%M"))

        # Linha 3
        tk.Label(frame_top, text="Repetir a cada:").grid(row=3, column=0, sticky="w", pady=10)
        
        frame_freq = tk.Frame(frame_top)
        frame_freq.grid(row=3, column=1, columnspan=3, sticky="w")
        
        self.entry_intervalo = tk.Entry(frame_freq, width=10)
        self.entry_intervalo.pack(side="left", padx=5)
        self.entry_intervalo.insert(0, "24")
        
        self.combo_unidade = ttk.Combobox(frame_freq, values=["Horas", "Minutos", "Dias"], state="readonly", width=10)
        self.combo_unidade.current(0)
        self.combo_unidade.pack(side="left", padx=5)
        
        tk.Label(frame_freq, text="(0 = Execução Única)", fg="gray").pack(side="left", padx=10)

        tk.Button(frame_top, text="AGENDAR TAREFA", command=self.adicionar_tarefa, bg="#ccffcc", height=2).grid(row=4, column=0, columnspan=4, sticky="we", pady=10)

        # --- 2. LISTAGEM ---
        frame_lista = tk.LabelFrame(root, text="Monitoramento de Tarefas", padx=10, pady=10)
        frame_lista.pack(fill="both", expand=True, padx=10, pady=5)
        
        cols = ("nome", "ult_exec", "prox_exec", "intervalo", "path")
        self.tree = ttk.Treeview(frame_lista, columns=cols, show="headings")
        
        self.tree.heading("nome", text="Nome")
        self.tree.heading("ult_exec", text="Última Execução")
        self.tree.heading("prox_exec", text="Próxima Execução")
        self.tree.heading("intervalo", text="Regra")
        self.tree.heading("path", text="Caminho")
        
        self.tree.column("nome", width=150)
        self.tree.column("ult_exec", width=140)
        self.tree.column("prox_exec", width=140)
        self.tree.column("intervalo", width=120)
        
        self.tree.pack(fill="both", expand=True)
        
        frame_botoes = tk.Frame(frame_lista)
        frame_botoes.pack(pady=5)
        tk.Button(frame_botoes, text="Excluir Tarefa", command=self.remover_tarefa, bg="#ffcccc").pack(side="left", padx=5)
        tk.Button(frame_botoes, text="Forçar Execução Agora", command=self.forcar_execucao, bg="#ccccff").pack(side="left", padx=5)

        # Inicialização
        self.carregar_dados()
        self.atualizar_visual() # <--- AQUI ESTÁ A CORREÇÃO
        self.iniciar_motor()

    # --- FUNÇÕES ---

    def buscar_arquivo(self):
        f = filedialog.askopenfilename(filetypes=[("Executáveis", "*.exe;*.bat;*.cmd;*.py"), ("Todos", "*.*")])
        if f:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, f)

    def adicionar_tarefa(self):
        nome = self.entry_nome.get()
        path = self.entry_path.get()
        data = self.entry_data.get()
        hora = self.entry_hora.get()
        interv_val = self.entry_intervalo.get()
        interv_uni = self.combo_unidade.get()

        if not (nome and path and data and hora and interv_val):
            messagebox.showwarning("Erro", "Preencha todos os campos.")
            return

        try:
            datetime.strptime(f"{data} {hora}", "%d/%m/%Y %H:%M")
            int(interv_val)
        except:
            messagebox.showerror("Erro", "Data inválida ou intervalo não numérico.")
            return

        nova = {
            "nome": nome,
            "path": path,
            "anchor_str": f"{data} {hora}",
            "interval_val": int(interv_val),
            "interval_unit": interv_uni,
            "last_run": "Nunca" # Campo novo para guardar histórico
        }
        
        self.tarefas.append(nova)
        self.salvar_dados()
        self.atualizar_visual()

    def remover_tarefa(self):
        sel = self.tree.selection()
        if not sel: return
        item = self.tree.item(sel)['values']
        self.tarefas = [t for t in self.tarefas if t['nome'] != item[0]]
        self.salvar_dados()
        self.atualizar_visual()

    def calcular_proxima(self, tarefa):
        agora = datetime.now()
        ancora = datetime.strptime(tarefa['anchor_str'], "%d/%m/%Y %H:%M")
        valor = tarefa['interval_val']
        unidade = tarefa['interval_unit']

        if valor == 0:
            return ancora if ancora > agora else None

        if unidade == "Minutos": delta = timedelta(minutes=valor)
        elif unidade == "Horas": delta = timedelta(hours=valor)
        elif unidade == "Dias": delta = timedelta(days=valor)
        
        proxima = ancora
        while proxima <= agora:
            proxima += delta
        return proxima

    def atualizar_visual(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        
        for t in self.tarefas:
            prox = self.calcular_proxima(t)
            prox_str = prox.strftime("%d/%m/%Y %H:%M:%S") if prox else "Concluído"
            regra = f"Cada {t['interval_val']} {t['interval_unit']}"
            # Pega o histórico, se não existir (tarefas antigas) mostra "Nunca"
            ult_exec = t.get('last_run', 'Nunca')
            
            self.tree.insert("", tk.END, values=(t['nome'], ult_exec, prox_str, regra, t['path']))

    def executar_processo(self, path, nome_tarefa=None):
        # 1. Atualiza o Horário da Última Execução na memória e JSON
        agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        if nome_tarefa:
            for t in self.tarefas:
                if t['nome'] == nome_tarefa:
                    t['last_run'] = agora_str
                    break
            self.salvar_dados()
            # Atualiza a interface visualmente (via thread segura) se necessário
            # self.root.after(0, self.atualizar_visual) -> Já é chamado no loop principal

        # 2. Log em Texto
        with open(ARQUIVO_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{agora_str}] Iniciando: {path}\n")
            
        # 3. Executa
        try:
            pasta = os.path.dirname(path)
            nome_arq = os.path.basename(path)
            cmd = f'start "Executando: {nome_arq}" "{path}"'
            subprocess.Popen(cmd, cwd=pasta, shell=True)
        except Exception as e:
            with open(ARQUIVO_LOG, "a", encoding="utf-8") as f:
                f.write(f"[{agora_str}] ERRO: {e}\n")

    def forcar_execucao(self):
        sel = self.tree.selection()
        if sel:
            valores = self.tree.item(sel)['values']
            nome = valores[0]
            path = valores[4] # O indice mudou porque adicionamos uma coluna
            
            self.executar_processo(path, nome_tarefa=nome)
            self.atualizar_visual() # Força atualização imediata da tela

    # --- PERSISTÊNCIA ---
    def salvar_dados(self):
        with open(ARQUIVO_DB, "w", encoding="utf-8") as f:
            json.dump(self.tarefas, f, indent=4)

    def carregar_dados(self):
        if os.path.exists(ARQUIVO_DB):
            try:
                with open(ARQUIVO_DB, "r", encoding="utf-8") as f:
                    self.tarefas = json.load(f)
            except: self.tarefas = []

    # --- MOTOR DO TEMPO ---
    def motor_loop(self):
        while True:
            agora = datetime.now()
            
            for t in self.tarefas:
                prox = self.calcular_proxima(t)
                
                if prox:
                    diff = (prox - agora).total_seconds()
                    # Margem de execução (1.5 segundos)
                    if 0 <= diff < 1.5:
                        # Chama a execução passando o Nome para salvar o histórico
                        self.root.after(0, lambda p=t['path'], n=t['nome']: self.executar_processo(p, n))
                        self.root.after(100, self.atualizar_visual)
                        time.sleep(1.5)

            time.sleep(1)

    def iniciar_motor(self):
        t = threading.Thread(target=self.motor_loop)
        t.daemon = True
        t.start()

if __name__ == "__main__":
    root = tk.Tk()
    app = AgendadorApp(root)
    root.mainloop()