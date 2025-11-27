# 📅 Agendador de Tarefas Master (Python)

Uma aplicação Desktop robusta para agendamento e automação de executáveis no Windows. Desenvolvido para oferecer controle preciso sobre intervalos de tempo e ancoragem de horários, superando limitações de agendadores comuns.

## ✨ Funcionalidades Principais

- **Interface Gráfica (GUI):** Interface limpa e nativa (Tkinter) para gerenciar tarefas sem mexer em código.
- **Ancoragem de Tempo (Time Anchoring):** Defina uma data e hora de início (ex: 25/11 às 14:00) e um intervalo (ex: 24 horas). O sistema garante que a execução ocorra sempre às 14:00, independente de quando o computador foi ligado.
- **Histórico de Execução:** Visualização na tabela de quando foi a última vez que a tarefa rodou.
- **Edição de Tarefas:** Permite alterar horários e caminhos de tarefas já cadastradas.
- **Execução Visível:** Os programas agendados abrem uma janela do CMD identificada, permitindo monitorar o status e logs do script em tempo real.
- **Logs de Auditoria:** Gera um arquivo `log_execucao.txt` registrando todas as tentativas de execução e erros.
- **Portabilidade:** Salva os dados em JSON na própria pasta, facilitando o transporte via Pen Drive ou rede.

## 🛠️ Tecnologias Utilizadas

O projeto foi construído utilizando **apenas bibliotecas nativas do Python**, garantindo máxima compatibilidade e leveza:

- `tkinter` (Interface Gráfica)
- `datetime` & `timedelta` (Cálculos temporais precisos)
- `subprocess` (Gerenciamento de processos do Windows)
- `threading` (Execução paralela para não travar a interface)
- `json` (Banco de dados local)

**Não é necessário instalar bibliotecas externas (como pandas ou schedule) para rodar o código fonte.**

---

## 🚀 Como Rodar (Código Fonte)

### Pré-requisitos
- Python 3.x instalado no Windows.

### Passo a Passo
1. Clone este repositório:
   ```bash
   git clone https://github.com/giggf/agendador-tarefas-python
   ```
2. Entre na pasta do projeto :
   ```bash
   cd "nome do repositório"
   ```
3. Execute o programa.

---

# 📦 Como criar um exectável para a aplicação (.exe)

Para tranforPara transformar este script em um software standalone (que funciona em computadores sem Python), utilizamos o **PyInstaller**.

## 1. Instalar o PyInstaller
   ```bash
   pip install pyinstaller
   ```
## 2. Gerar o executável
   ```bash
   pyinstaller --noconsole --onefile --clean agendador.pyw
   ```
## 3. Onde está o arquivo?
O executável final estará na pasta `dist`.

---

# ⚠️ Solução de Problemas (Windows Long Path)

Se ao tentar criar o executável você receber um erro vermelho mencionando **"Windows Long Path support"** ou **"file does not exist"**, siga estes passos (comuns em PCs corporativos):

1. Crie uma pasta na raiz do disco C, exemplo: C:\Dev.
2. Copie o arquivo agendador.pyw para lá.
3. Abra o terminal nessa pasta (cd C:\Dev).
4. Crie um ambiente virtual curto: 
   ```bash
   python -m venv venv
   .\venv\Scripts\activate
   pip install pyinstaller
   ```
5. Gere o executável.

---

# 📂 Estrutura de Arquivos
Ao rodar, o programa gerará automaticamente dois arquivos na mesma pasta:`
- `tarefas.json`: Banco de dados das suas tarefas.
- `log_execucao.txt`: Histórico de erros e sucessos.

**Nota:** Mantenha o tarefas.json junto do .exe se mover o programa de lugar, para não perder seus agendamentos.



