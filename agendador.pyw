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
        self.root.title("Agendador Master 3.6")
        self.root.geometry("1000x650")
        
        self.tarefas = []
        self.tarefa_em_edicao_index = None
        
        # --- VINCULA A TECLA F5 ---
        self.root.bind('<F5>', self.atualizar_tudo)
        
        # --- 1. ÁREA DE CADASTRO ---
        self.frame_top = tk.LabelFrame(root, text="Configurar / Editar Tarefa", padx=15, pady=15)
        self.frame_top.pack(fill="x", padx=10, pady=5)
        
        # Linha 1
        tk.Label(self.frame_top, text="Nome da Tarefa:").grid(row=0, column=0, sticky="w")
        self.entry_nome = tk.Entry(self.frame_top, width=30)
        self.entry_nome.grid(row=0, column=1, padx=5, sticky="w")
        
        tk.Button(self.frame_top, text="Selecionar Executável...", command=self.buscar_arquivo).grid(row=0, column=2, padx=5)
        self.entry_path = tk.Entry(self.frame_top, width=40)
        self.entry_path.grid(row=0, column=3, padx=5)

        ttk.Separator(self.frame_top, orient='horizontal').grid(row=1, column=0, columnspan=4, sticky="ew", pady=15)

        # Linha 2
        tk.Label(self.frame_top, text="Começar dia (DD/MM/AAAA):").grid(row=2, column=0, sticky="w")
        self.entry_data = tk.Entry(self.frame_top, width=15)
        self.entry_data.grid(row=2, column=1, sticky="w", padx=5)
        self.entry_data.insert(0, datetime.now().strftime("%d/%m/%Y"))

        tk.Label(self.frame_top, text="Hora (HH:MM):").grid(row=2, column=2, sticky="e")
        self.entry_hora = tk.Entry(self.frame_top, width=10)
        self.entry_hora.grid(row=2, column=3, sticky="w", padx=5)
        self.entry_hora.insert(0, datetime.now().strftime("%H:%M"))

        # Linha 3
        tk.Label(self.frame_top, text="Repetir a cada:").grid(row=3, column=0, sticky="w", pady=10)
        
        frame_freq = tk.Frame(self.frame_top)
        frame_freq.grid(row=3, column=1, columnspan=3, sticky="w")
        
        self.entry_intervalo = tk.Entry(frame_freq, width=10)
        self.entry_intervalo.pack(side="left", padx=5)
        self.entry_intervalo.insert(0, "24")
        
        self.combo_unidade = ttk.Combobox(frame_freq, values=["Horas", "Minutos", "Dias"], state="readonly", width=10)
        self.combo_unidade.current(0)
        self.combo_unidade.pack(side="left", padx=5)
        
        tk.Label(frame_freq, text="(0 = Execução Única)", fg="gray").pack(side="left", padx=10)

        # Botões de Ação
        frame_action = tk.Frame(self.frame_top)
        frame_action.grid(row=4, column=0, columnspan=4, sticky="we", pady=10)
        
        self.btn_salvar = tk.Button(frame_action, text="AGENDAR TAREFA", command=self.salvar_ou_adicionar, bg="#ccffcc", height=2)
        self.btn_salvar.pack(side="left", fill="x", expand=True, padx=5)
        
        self.btn_cancelar = tk.Button(frame_action, text="Cancelar Edição", command=self.cancelar_edicao, bg="#f0f0f0", height=2)
        self.btn_cancelar.pack(side="left", padx=5)
        self.btn_cancelar.config(state="disabled")

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
        
        # --- BOTÕES INFERIORES ---
        frame_botoes = tk.Frame(frame_lista)
        frame_botoes.pack(pady=5)
        
        # Botão Atualizar (Sem feedback visual no título)
        tk.Button(frame_botoes, text="🔄 Atualizar Lista (F5)", command=self.atualizar_tudo, bg="#e6e6e6").pack(side="left", padx=5)
        
        tk.Button(frame_botoes, text="Editar Selecionada", command=self.preparar_edicao, bg="#fffccc").pack(side="left", padx=5)
        tk.Button(frame_botoes, text="Excluir Tarefa", command=self.remover_tarefa, bg="#ffcccc").pack(side="left", padx=5)
        tk.Button(frame_botoes, text="Forçar Execução Agora", command=self.forcar_execucao, bg="#ccccff").pack(side="left", padx=5)

        # Inicialização
        self.carregar_dados()
        self.atualizar_visual()
        self.iniciar_motor()

    # --- FUNÇÕES VISUAIS ---

    def buscar_arquivo(self):
        f = filedialog.askopenfilename(filetypes=[("Executáveis", "*.exe;*.bat;*.cmd;*.py"), ("Todos", "*.*")])
        if f:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, f)
            
    # --- FUNÇÃO DE REFRESH LIMPA ---
    def atualizar_tudo(self, event=None):
        # Apenas recarrega e desenha, sem mudar título
        self.carregar_dados()
        self.atualizar_visual()

    def cancelar_edicao(self):
        self.entry_nome.delete(0, tk.END)
        self.entry_path.delete(0, tk.END)
        self.entry_intervalo.delete(0, tk.END)
        self.entry_intervalo.insert(0, "24")
        
        self.entry_data.delete(0, tk.END)
        self.entry_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.entry_hora.delete(0, tk.END)
        self.entry_hora.insert(0, datetime.now().strftime("%H:%M"))
        
        self.tarefa_em_edicao_index = None
        self.btn_salvar.config(text="AGENDAR TAREFA", bg="#ccffcc")
        self.frame_top.config(text="Configurar Nova Tarefa")
        self.btn_cancelar.config(state="disabled")

    def preparar_edicao(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma tarefa na lista para editar.")
            return

        item = self.tree.item(sel)['values']
        nome_alvo = item[0]
        
        index = -1
        tarefa_encontrada = None
        for i, t in enumerate(self.tarefas):
            if t['nome'] == nome_alvo:
                index = i
                tarefa_encontrada = t
                break
        
        if index != -1:
            self.tarefa_em_edicao_index = index
            t = tarefa_encontrada
            
            self.entry_nome.delete(0, tk.END)
            self.entry_nome.insert(0, t['nome'])
            
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, t['path'])
            
            self.entry_intervalo.delete(0, tk.END)
            self.entry_intervalo.insert(0, t['interval_val'])
            
            self.combo_unidade.set(t['interval_unit'])
            
            try:
                data_hora = t['anchor_str'].split(' ')
                self.entry_data.delete(0, tk.END)
                self.entry_data.insert(0, data_hora[0])
                self.entry_hora.delete(0, tk.END)
                self.entry_hora.insert(0, data_hora[1])
            except: pass
            
            self.btn_salvar.config(text="SALVAR ALTERAÇÕES", bg="#ffd700")
            self.frame_top.config(text=f"Editando: {t['nome']}")
            self.btn_cancelar.config(state="normal")

    def salvar_ou_adicionar(self):
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

        nova_tarefa = {
            "nome": nome,
            "path": path,
            "anchor_str": f"{data} {hora}",
            "interval_val": int(interv_val),
            "interval_unit": interv_uni,
            "last_run": "Nunca" 
        }

        if self.tarefa_em_edicao_index is not None:
            hist_antigo = self.tarefas[self.tarefa_em_edicao_index].get('last_run', 'Nunca')
            nova_tarefa['last_run'] = hist_antigo
            self.tarefas[self.tarefa_em_edicao_index] = nova_tarefa
            messagebox.showinfo("Sucesso", "Tarefa atualizada com sucesso!")
        else:
            self.tarefas.append(nova_tarefa)
        
        self.salvar_dados()
        self.atualizar_visual()
        self.cancelar_edicao()

    def remover_tarefa(self):
        sel = self.tree.selection()
        if not sel: return
        item = self.tree.item(sel)['values']
        self.tarefas = [t for t in self.tarefas if t['nome'] != item[0]]
        self.salvar_dados()
        self.atualizar_visual()

    # --- LÓGICA DO MOTOR ---

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
            ult_exec = t.get('last_run', 'Nunca')
            
            self.tree.insert("", tk.END, values=(t['nome'], ult_exec, prox_str, regra, t['path']))

    def executar_processo(self, path, nome_tarefa=None):
        agora_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        if nome_tarefa:
            for t in self.tarefas:
                if t['nome'] == nome_tarefa:
                    t['last_run'] = agora_str
                    break
            self.salvar_dados()

        with open(ARQUIVO_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{agora_str}] Iniciando: {path}\n")
            
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
            path = valores[4]
            self.executar_processo(path, nome_tarefa=nome)
            self.atualizar_visual()

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

    def motor_loop(self):
        while True:
            agora = datetime.now()
            for t in self.tarefas:
                prox = self.calcular_proxima(t)
                if prox:
                    diff = (prox - agora).total_seconds()
                    if 0 <= diff < 1.5:
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