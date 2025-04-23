import pandas as pd
import os
import re
import datetime
import pytz
import logging

# ------------------------ Configurações Iniciais ------------------------

ARQUIVO_EXCEL = "MotoristasRegistro.xlsx"
ARQUIVO_LOG = 'registro.log'
CONTADOR_PATH = 'contador.txt'

fuso_brasil = pytz.timezone('America/Sao_Paulo')

logging.basicConfig(
    filename=ARQUIVO_LOG, 
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

conversas_ativas = {}

# ------------------------ Frases de Incentivo ------------------------

marcos = {
    3: "🔥 Uma Maquina! Já são 3 cadastros!",
    5: "🚀 Já tá voando, hein! 5 cadastros na conta!",
    10: "👑 O REI DO CADASTRO ESTÁ ONLINE! Dez cadastros!",
    20: "📈 RH vai te contratar só pra cadastrar!",
    30: "🏆 Isso aqui virou maratona de cadastro? 30 já!",
    50: "💥 Luiz, para de humilhar! 50 motoristas?!"
}

def carregar_contador():
    if os.path.exists(CONTADOR_PATH):
        with open(CONTADOR_PATH, 'r') as f:
            return int(f.read().strip())
    return 0

def salvar_contador(contador):
    with open(CONTADOR_PATH, 'w') as f:
        f.write(str(contador))

def atualizar_contador():
    contador = carregar_contador() + 1
    salvar_contador(contador)
    if contador in marcos:
        print(f"\n🎉 {marcos[contador]}\n")
    return contador

# ------------------------ Funções de Identificação e Extração ------------------------

def identificar_tipo(texto):
    texto = texto.lower()
    if "tac" in texto or "veículo da empresa" in texto:
        return "TAC"
    elif "agregado" in texto or "veículo próprio" in texto:
        return "Agregado"
    return None

def pegar_dados_da_mensagem(mensagem):
    padroes = {
        "Nome": r"(?i)(?:meu nome é|sou o|sou a|me chamo|nome:)\s*([A-ZÁ-Úa-zá-ú\s]+?)(?=\s*(?:cpf|telefone|$))",
        "CPF": r"\d{3}\.?\d{3}\.?\d{3}-?\d{2}",
        "Telefone": r"(?:\+?55)?\s*\(?\d{2}\)?\s*\d{4,5}-?\d{4}",
        "Cidade": r"(?i)(?:cidade|moro em|base:)\s*([A-ZÁ-Úa-zá-ú\s]+?)(?=\s*(?:cpf|telefone|$))",
        "Placa": r"[A-Z]{3}\d[A-Z0-9]\d{2}",
        "Curso": r"(?i)curso concluído\??[:\-]?\s*(sim|não)"
    }
    dados = {campo: re.search(pad, mensagem) for campo, pad in padroes.items()}
    return {
        "Nome": dados["Nome"].group(1).strip() if dados["Nome"] else "",
        "CPF": re.sub(r"\D", "", dados["CPF"].group()) if dados["CPF"] else "",
        "Telefone": re.sub(r"\D", "", dados["Telefone"].group()) if dados["Telefone"] else "",
        "Cidade": dados["Cidade"].group(1).strip() if dados["Cidade"] else "",
        "Tipo": identificar_tipo(mensagem),
        "Placa": dados["Placa"].group(0) if dados["Placa"] else "",
        "Curso": dados["Curso"].group(1).capitalize() if dados["Curso"] else ""
    }

# ------------------------ Verificação e Atualização ------------------------

def ta_tudo_preenchido(dados):
    if dados['Tipo'] == "TAC":
        return all([dados["Nome"], dados["CPF"], dados["Telefone"], dados["Cidade"], dados["Curso"]])
    elif dados['Tipo'] == "Agregado":
        return all([dados["Nome"], dados["CPF"], dados["Telefone"], dados["Cidade"], dados["Placa"]])
    return False

def atualiza_conversa_com_motorista(mensagem):
    dados = pegar_dados_da_mensagem(mensagem)
    telefone = dados['Telefone']
    if not telefone:
        print("⚠️ Opa! Não achei o telefone. Tenta de novo, por favor.")
        return

    if telefone not in conversas_ativas:
        conversas_ativas[telefone] = dados
    else:
        for chave in dados:
            if dados[chave]:
                conversas_ativas[telefone][chave] = dados[chave]

    if ta_tudo_preenchido(conversas_ativas[telefone]):
        salvar_no_excel(conversas_ativas[telefone])
        atualizar_contador()
        del conversas_ativas[telefone]
        print("✅ Cadastro salvo com sucesso!")
    else:
        print("📌 Pendente. Me manda o resto quando puder!")

# ------------------------ Salvamento no Excel ------------------------

def salvar_no_excel(dados):
    tipo = dados["Tipo"]
    aba = tipo if tipo in ["TAC", "Agregado"] else "Outros"

    if not os.path.exists(ARQUIVO_EXCEL):
        with pd.ExcelWriter(ARQUIVO_EXCEL, engine='openpyxl') as writer:
            for aba_nome in ["TAC", "Agregado", "Contatos Incompletos"]:
                pd.DataFrame().to_excel(writer, sheet_name=aba_nome, index=False)

    aba_certa = aba if ta_tudo_preenchido(dados) else "Contatos Incompletos"
    df = pd.read_excel(ARQUIVO_EXCEL, sheet_name=aba_certa)
    registro = dados.copy()
    registro["DataCadastro"] = datetime.datetime.now(fuso_brasil).strftime('%d/%m/%Y %H:%M')
    registro["Status"] = "Completo" if ta_tudo_preenchido(dados) else "Em andamento"

    df = pd.concat([df, pd.DataFrame([registro])], ignore_index=True)
    with pd.ExcelWriter(ARQUIVO_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=aba_certa, index=False)

# ------------------------ Interface Menu ------------------------

def mostrar_modelos_exemplo():
    total = carregar_contador()
    proximo = next((m for m in sorted(marcos) if m > total), None)
    print(f"\n📊 Total de cadastros completos hoje: {total}")
    if proximo:
        print(f"🎯 Faltam {proximo - total} para: {marcos[proximo]}")
    else:
        print("🏅 Você já bateu todos os marcos do dia! 👏")

    print("\nModelo Agregado:")
    print("""
📝 *Cadastro de Motorista*
👤 Nome: 
🪪 CPF: 
📱 Telefone: 
🏛️ Cidade: 
🚚 Modalidade: Agregado (veículo próprio)
🚘 Placa: 
⚠️ Encaminhe essa mensagem no terminal para registrar no sistema.
    """)
    print("\nModelo TAC:")
    print("""
📝 *Cadastro de Motorista*
👤 Nome: 
🪪 CPF: 
📱 Telefone: 
🏛️ Cidade: 
🚚 Modalidade: TAC (veículo da empresa)
🎓 Curso concluído?(SIM/NÃO):
⚠️ Encaminhe essa mensagem no terminal para registrar no sistema.
    """)

def entrada_interativa_manual():
    print("\nVamos preencher os dados manualmente:")
    dados = {}
    dados["Nome"] = input("🔀 Nome: ")
    dados["CPF"] = re.sub(r"\D", "", input("🔀 CPF: "))
    dados["Telefone"] = re.sub(r"\D", "", input("🔀 Telefone: "))
    dados["Cidade"] = input("🔀 Cidade: ")
    tipo = input("🔀 Modalidade (TAC ou Agregado): ").strip().upper()
    dados["Tipo"] = tipo

    if tipo == "TAC":
        dados["Curso"] = input("🔀 Curso concluído? (Sim/Não): ").strip().capitalize()
        dados["Placa"] = ""
    elif tipo == "AGREGADO":
        dados["Placa"] = input("🔀 Placa: ").strip().upper()
        dados["Curso"] = ""
    else:
        print("⚠️ Modalidade inválida. Tenta de novo.")
        return

    if ta_tudo_preenchido(dados):
        salvar_no_excel(dados)
        atualizar_contador()
        print("✅ Cadastro completo e salvo com sucesso!")
    else:
        salvar_no_excel(dados)
        print("📌 Anotei o que consegui. Me manda o que faltar depois, beleza?")


### 📑 **Relatório Diário**

import datetime

def gerar_relatorio_diario():
    # Carregar o arquivo Excel e ler as abas TAC e Agregado
    df_tac = pd.read_excel(ARQUIVO_EXCEL, sheet_name="TAC")
    df_agregado = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Agregado")
    
    # Filtrar os registros completos
    completos_tac = df_tac[df_tac['Status'] == 'Completo']
    completos_agregado = df_agregado[df_agregado['Status'] == 'Completo']
    
    # Filtrar os registros incompletos
    incompletos_tac = df_tac[df_tac['Status'] != 'Completo']
    incompletos_agregado = df_agregado[df_agregado['Status'] != 'Completo']
    
    # Contagem de motoristas
    total_completos = len(completos_tac) + len(completos_agregado)
    total_incompletos = len(incompletos_tac) + len(incompletos_agregado)
    total_motoristas = total_completos + total_incompletos
    
    # Gerar relatório
    data_hoje = datetime.datetime.now(fuso_brasil).strftime("%d/%m/%Y")
    relatorio = f"""
    Relatório Diário - {data_hoje}

    🏅 Total de motoristas cadastrados: {total_motoristas}
    ✅ Total de cadastros completos: {total_completos}
    ❌ Total de cadastros incompletos: {total_incompletos}

    🔥 Status de cadastros:
    - TAC: {len(completos_tac)} completos, {len(incompletos_tac)} incompletos
    - Agregado: {len(completos_agregado)} completos, {len(incompletos_agregado)} incompletos

    📝 Cada cadastro completo gera um marco de progresso!

    Obrigado por continuar com o sistema! 🚛
    """

    # Salvar o relatório em um arquivo .txt
    with open(f"relatorio_diario_{data_hoje}.txt", "w", encoding="utf-8") as f:
        f.write(relatorio)
    
    print("📄 Relatório diário gerado com sucesso!")

# Adicionar a opção 6 no menu para gerar o relatório diário
def menu_principal():
    print("\n" + "=" * 40)
    print(" 🚛 SISTEMA DE MOTORISTAS BLD ")
    print("=" * 40)
    while True:
        print("""
MENU PRINCIPAL:
1️⃣ Inserir dados manualmente (um por vez)
2️⃣ Inserir dados colando mensagem pronta
3️⃣ Ver progresso e modelos de mensagens
6️⃣ Gerar relatório diário
7️⃣ Sair
        """)
        opcao = input("Escolha uma opção (1-7): ").strip()
        if opcao == "1":
            entrada_interativa_manual()
        elif opcao == "2":
            mensagem = input("\nCole a mensagem do motorista:\n> ")
            atualiza_conversa_com_motorista(mensagem)
        elif opcao == "3":
            mostrar_modelos_exemplo()
        elif opcao == "6":
            gerar_relatorio_diario()
        elif opcao == "7":
            print("\nValeu por usar o sistema! Té mais 🚀")
            break
        else:
            print("⚠️ Opção inválida. Tenta de novo!")

if __name__ == "__main__":
    menu_principal()
