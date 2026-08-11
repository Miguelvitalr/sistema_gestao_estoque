import sqlite3
import hashlib
import smtplib
from email.mime.text import MIMEText
from datetime import datetime
import openpyxl
from openpyxl.chart import BarChart, Reference

ESTOQUE_ALERTA_EMAIL = 10

# --- Configuracao de e-mail ---
# Use uma "senha de app" do Gmail, nao a senha normal da conta.
# Como gerar: Conta Google > Seguranca > Verificacao em duas etapas > Senhas de app
EMAIL_REMETENTE = "sistemadeestoque67@gmail.com"
EMAIL_SENHA = "bout auuq ydro zzdx"
SMTP_SERVIDOR = "smtp.gmail.com"
SMTP_PORTA = 587


def enviar_alerta_estoque(destinatario, produto, quantidade):
    assunto = f"Alerta de estoque baixo: {produto}"
    corpo = (
        f"O produto '{produto}' esta com estoque baixo.\n"
        f"Quantidade atual: {quantidade} unidades.\n\n"
        f"Sistema de Gestao de Estoque"
    )

    msg = MIMEText(corpo)
    msg["Subject"] = assunto
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = destinatario

    try:
        servidor = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA)
        servidor.starttls()
        servidor.login(EMAIL_REMETENTE, EMAIL_SENHA)
        servidor.sendmail(EMAIL_REMETENTE, destinatario, msg.as_string())
        servidor.quit()
        print(f"E-mail de alerta enviado para {destinatario}")
    except Exception as erro:
        print(f"Erro ao enviar e-mail: {erro}")


def verificar_e_alertar(gestao, produto, email_destino):
    quantidade_atual = gestao.consultar_estoque(produto)
    if 0 < quantidade_atual <= ESTOQUE_ALERTA_EMAIL:
        enviar_alerta_estoque(email_destino, produto, quantidade_atual)

class Gestao: # planta casa 

    def __init__(self, banco): # constrututor 
        self.conn = sqlite3.connect(banco)
        self.criar_tabela_estoque()
        self.criar_tabela_fornecedores()

    def criar_tabela_estoque(self): # criacao tabela 
        cursor = self.conn.cursor() # conectar cursor , info vai pro banco 
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS estoque (
        id INTEGER PRIMARY KEY,
        produto TEXT UNIQUE,
        quantidade INTEGER,
        preco REAL
        )
        ''')
        self.conn.commit()

    def criar_tabela_fornecedores(self): # criacao tabela de fornecedores
        cursor = self.conn.cursor()
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS fornecedores (
        id INTEGER PRIMARY KEY,
        nome TEXT,
        celular TEXT,
        produto TEXT
        )
        ''')
        self.conn.commit()

    def adicionar_produto(self, produto, quantidade, preco=None): # parametros que adicionar prod recebe 
        cursor = self.conn.cursor() # conectar com cursor que escreve no banco 
        cursor.execute("SELECT quantidade, preco FROM estoque WHERE produto=?", (produto,)) # parametro e oque sera substituido depois 
        resultado = cursor.fetchone() # nao sei ainda 
        if resultado:
            nova_quantidade = resultado[0] + quantidade # soma quantidade na posicao 1 com a nova quantidade adicionada 
            preco_final = preco if preco is not None else resultado[1] 
            cursor.execute("UPDATE estoque SET quantidade=?, preco=? WHERE produto=?",
                            (nova_quantidade, preco_final, produto))
        else:
            if preco is None:
                preco = 0
            nova_quantidade = quantidade
            cursor.execute("INSERT INTO estoque (produto, quantidade, preco) VALUES (?,?,?)",
                            (produto, quantidade, preco))

        self.conn.commit()

        if nova_quantidade <= ESTOQUE_ALERTA_EMAIL:
            print(f"AVISO: estoque de {produto} esta baixo ({nova_quantidade} unidades)")

    def consultar_produto_por_id(self, produto_id):
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, produto, quantidade, preco FROM estoque WHERE id=?", (produto_id,))
        return cursor.fetchone()

    def adicionar_produto_por_id(self, produto_id, quantidade, preco=None):
        produto_existente = self.consultar_produto_por_id(produto_id)

        if not produto_existente:
            print(f"Nenhum produto encontrado com o id {produto_id}")
            return False

        _, produto, quantidade_atual, preco_atual = produto_existente
        nova_quantidade = quantidade_atual + quantidade
        preco_final = preco if preco is not None else preco_atual

        cursor = self.conn.cursor()
        cursor.execute("UPDATE estoque SET quantidade=?, preco=? WHERE id=?",
                        (nova_quantidade, preco_final, produto_id))
        self.conn.commit()

        print(f"{produto} (id {produto_id}) atualizado: nova quantidade = {nova_quantidade}")

        if nova_quantidade <= ESTOQUE_ALERTA_EMAIL:
            print(f"AVISO: estoque de {produto} esta baixo ({nova_quantidade} unidades)")

        return True

    def executar_venda(self, produto, quantidade):
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT quantidade FROM estoque WHERE produto=?", (produto,))
        resultado = cursor.fetchone()
        if resultado:
            estoque_atual = resultado[0]
            if estoque_atual >= quantidade:
                nova_quantidade = estoque_atual - quantidade
                cursor.execute("UPDATE estoque SET quantidade=? WHERE produto=? ",
                                (nova_quantidade, produto))
                self.conn.commit()

                if nova_quantidade <= ESTOQUE_ALERTA_EMAIL:
                    print(f"AVISO: estoque de {produto} esta baixo ({nova_quantidade} unidades)")

                return True
            else:
                print(f"estoque insuficiente para a retirada : {produto}")
                return False
        else:
            print(f"{produto} nao encontrada no estoque ")
            return False

    def remover_produto(self, produto):
        cursor = self.conn.cursor()
        cursor.execute("SELECT produto FROM estoque WHERE produto=?", (produto,))
        resultado = cursor.fetchone()

        if not resultado:
            print(f"{produto} nao encontrado no estoque ")
            return False

        cursor.execute("DELETE FROM estoque WHERE produto=?", (produto,))
        self.conn.commit()
        print(f"{produto} removido do estoque")
        return True

    def consultar_estoque(self, produto):
        cursor = self.conn.cursor()
        cursor.execute("SELECT quantidade FROM estoque WHERE produto=?", (produto,))
        resultado = cursor.fetchone()
        if resultado:
            return resultado[0]
        else:
            return 0

    def consultar_preco(self, produto):
        cursor = self.conn.cursor()
        cursor.execute("SELECT preco FROM estoque WHERE produto=?", (produto,))
        resultado = cursor.fetchone()
        if resultado:
            return resultado[0]
        else:
            return None

    def lista_estoque(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT produto, quantidade, preco FROM estoque WHERE quantidade > 0")
        produtos = cursor.fetchall()
        return produtos

    def valor_total_estoque(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT SUM(quantidade * preco) FROM estoque")
        resultado = cursor.fetchone()
        return resultado[0] if resultado[0] is not None else 0

    def gerar_grafico_estoque(self, caminho_saida="relatorio_estoque.xlsx"):
        produtos = self.lista_estoque()  # reaproveita o metodo que ja existe

        if not produtos:
            print("Nenhum produto em estoque para gerar o grafico.")
            return

        # --- Monta a planilha ---
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Estoque"

        ws.append(["Produto", "Quantidade", "Preco"])
        for produto, quantidade, preco in produtos:
            ws.append([produto, quantidade, preco])

        ultima_linha = ws.max_row  # calculado uma unica vez, guardado em variavel

        # --- Cria o grafico de barras (quantidade por produto) ---
        grafico = BarChart()
        grafico.type = "col"
        grafico.title = "Quantidade em Estoque por Produto"
        grafico.x_axis.title = "Produto"
        grafico.y_axis.title = "Quantidade"
        grafico.width = 20
        grafico.height = 12

        dados_ref = Reference(ws, min_col=2, min_row=1, max_row=ultima_linha)
        categorias_ref = Reference(ws, min_col=1, min_row=2, max_row=ultima_linha)

        grafico.add_data(dados_ref, titles_from_data=True)
        grafico.set_categories(categorias_ref)

        ws.add_chart(grafico, "E2")

        wb.save(caminho_saida)
        print(f"Grafico de estoque gerado: {caminho_saida}")

    def adicionar_fornecedor(self, nome, celular, produto): # cadastra um fornecedor novo
        cursor = self.conn.cursor()
        cursor.execute(
            "INSERT INTO fornecedores (nome, celular, produto) VALUES (?,?,?)",
            (nome, celular, produto)
        )
        self.conn.commit()
        print(f"fornecedor cadastrado: {nome} - {produto} - {celular}")

    def listar_fornecedores(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, nome, celular, produto FROM fornecedores")
        return cursor.fetchall()

    def consultar_fornecedores_por_produto(self, produto):
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, nome, celular, produto FROM fornecedores WHERE produto=?", (produto,))
        return cursor.fetchall()

    def remover_fornecedor(self, fornecedor_id):
        cursor = self.conn.cursor()
        cursor.execute("SELECT nome FROM fornecedores WHERE id=?", (fornecedor_id,))
        resultado = cursor.fetchone()

        if not resultado:
            print(f"fornecedor com id {fornecedor_id} nao encontrado ")
            return

        cursor.execute("DELETE FROM fornecedores WHERE id=?", (fornecedor_id,))
        self.conn.commit()
        print(f"fornecedor {resultado[0]} removido")


class Usuarios:

    def __init__(self, banco):
        self.conn = sqlite3.connect(banco)
        self.criar_tabela_usuarios()

    def criar_tabela_usuarios(self):
        cursor = self.conn.cursor()
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY,
        username TEXT UNIQUE,
        senha_hash TEXT,
        email TEXT
        )
        ''')
        self.conn.commit()

    def _gerar_hash(self, senha):
        return hashlib.sha256(senha.encode()).hexdigest()

    def cadastrar_usuario(self, username, senha, email):
        cursor = self.conn.cursor()
        cursor.execute("SELECT username FROM usuarios WHERE username=?", (username,))
        if cursor.fetchone():
            print(f"usuario '{username}' ja existe, escolha outro nome")
            return False

        senha_hash = self._gerar_hash(senha)
        cursor.execute(
            "INSERT INTO usuarios (username, senha_hash, email) VALUES (?,?,?)",
            (username, senha_hash, email)
        )
        self.conn.commit()
        print(f"usuario '{username}' cadastrado com sucesso")
        return True

    def login(self, username, senha):
        cursor = self.conn.cursor()
        senha_hash = self._gerar_hash(senha)
        cursor.execute(
            "SELECT id, username, email FROM usuarios WHERE username=? AND senha_hash=?",
            (username, senha_hash)
        )
        resultado = cursor.fetchone()
        if resultado:
            id_usuario, username, email = resultado
            print(f"login realizado com sucesso, bem-vindo {username}")
            return {"id": id_usuario, "username": username, "email": email}
        else:
            print("usuario ou senha incorretos")
            return None


class Vendas: # planta 

    def __init__(self, banco, gestao): 
        self.conn = sqlite3.connect(banco)
        self.gestao = gestao
        self.criar_tabela_vendas()

    def criar_tabela_vendas(self):
        cursor = self.conn.cursor()
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS vendas (
        id INTEGER PRIMARY KEY,
        produto TEXT,
        quantidade INTEGER,
        preco_unitario REAL,
        valor_total REAL,
        data TEXT
        )
        ''')
        self.conn.commit()

    def registrar_venda(self, produto, quantidade):
        preco = self.gestao.consultar_preco(produto) # chamar uma funcao dentro da outra classe 

        if preco is None:
            print(f"{produto} nao encontrado no estoque ")
            return

        estoque_atual = self.gestao.consultar_estoque(produto)
        if estoque_atual < quantidade:
            print(f"estoque insuficiente para a venda : {produto} (disponivel: {estoque_atual})")
            return

        sucesso = self.gestao.executar_venda(produto, quantidade)
        if not sucesso:
            return

        valor_total = preco * quantidade
        data = datetime.now().strftime("%Y-%m-%d %H:%M:%S") # info da venda (anos meses dias ) , aqui e usado datetime 

        cursor = self.conn.cursor()
        cursor.execute(
            "INSERT INTO vendas (produto, quantidade, preco_unitario, valor_total, data) VALUES (?,?,?,?,?)",
            (produto, quantidade, preco, valor_total, data)
        )
        self.conn.commit()
        print(f"venda registrada: {quantidade}x {produto} - total R${valor_total:.2f}")

    def remover_venda(self, venda_id):
        cursor = self.conn.cursor()
        cursor.execute("SELECT produto, quantidade FROM vendas WHERE id=?", (venda_id,))
        resultado = cursor.fetchone()

        if not resultado:
            print(f"venda com id {venda_id} nao encontrada ")
            return

        produto, quantidade = resultado

        self.gestao.adicionar_produto(produto, quantidade)

        cursor.execute("DELETE FROM vendas WHERE id=?", (venda_id,))
        self.conn.commit()
        print(f"venda {venda_id} removida, {quantidade}x {produto} devolvido ao estoque")

    def listar_vendas(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, produto, quantidade, preco_unitario, valor_total, data FROM vendas")
        return cursor.fetchall()


def exibir_menu():
    print("\n===== SISTEMA DE GESTAO DE ESTOQUE =====")
    print("1 - Adicionar produto ao estoque")
    print("2 - Registrar venda")
    print("3 - Remover produto do estoque")
    print("4 - Remover venda")
    print("5 - Adicionar fornecedor")
    print("6 - Remover fornecedor")
    print("7 - Listar estoque")
    print("8 - Listar fornecedores")
    print("9 - Listar vendas")
    print("10 - Ver valor total do estoque")
    print("11 - Adicionar quantidade a um produto pelo ID")
    print("0 - Sair e gerar grafico atualizado")
    print("=========================================")


def pedir_inteiro(mensagem):
    while True:
        try:
            return int(input(mensagem).strip())
        except ValueError:
            print("Valor invalido, digite um numero inteiro.")


def pedir_float(mensagem):
    while True:
        try:
            return float(input(mensagem).strip())
        except ValueError:
            print("Valor invalido, digite um numero (ex: 10.50).")


def tela_login(usuarios):
    while True:
        print("\n===== LOGIN =====")
        print("1 - Entrar")
        print("2 - Criar conta")
        print("0 - Sair do programa")
        escolha = input("Escolha uma opcao: ").strip()

        if escolha == "1":
            username = input("Usuario: ").strip()
            senha = input("Senha: ").strip()
            usuario_logado = usuarios.login(username, senha)
            if usuario_logado:
                return usuario_logado

        elif escolha == "2":
            username = input("Novo usuario: ").strip()
            senha = input("Nova senha: ").strip()
            email = input("E-mail (para receber alertas de estoque): ").strip()
            usuarios.cadastrar_usuario(username, senha, email)

        elif escolha == "0":
            exit()

        else:
            print("Opcao invalida, tente novamente.")


if __name__ == "__main__":
    gestao = Gestao("estoque.db")
    vendas = Vendas("estoque.db", gestao)
    usuarios = Usuarios("estoque.db")

    usuario_logado = tela_login(usuarios)
    email_usuario = usuario_logado["email"]

    while True:
        exibir_menu()
        opcao = input("Escolha uma opcao: ").strip()

        if opcao == "1":
            produto = input("Nome do produto: ").strip()
            quantidade = pedir_inteiro("Quantidade: ")
            preco = pedir_float("Preco (ex: 10.50): ")
            gestao.adicionar_produto(produto, quantidade, preco)

        elif opcao == "2":
            produto = input("Nome do produto: ").strip()
            quantidade = pedir_inteiro("Quantidade vendida: ")
            vendas.registrar_venda(produto, quantidade)
            verificar_e_alertar(gestao, produto, email_usuario)

        elif opcao == "3":
            produto = input("Nome do produto a remover: ").strip()
            gestao.remover_produto(produto)

        elif opcao == "4":
            venda_id = pedir_inteiro("ID da venda a remover: ")
            vendas.remover_venda(venda_id)

        elif opcao == "5":
            nome = input("Nome do fornecedor: ").strip()
            celular = input("Celular: ").strip()
            produto = input("Produto fornecido: ").strip()
            gestao.adicionar_fornecedor(nome, celular, produto)

        elif opcao == "6":
            fornecedor_id = pedir_inteiro("ID do fornecedor a remover: ")
            gestao.remover_fornecedor(fornecedor_id)

        elif opcao == "7":
            produtos = gestao.lista_estoque()
            if not produtos:
                print("Estoque vazio.")
            for produto, quantidade, preco in produtos:
                print(f"{produto} | quantidade: {quantidade} | preco: R${preco:.2f}")

        elif opcao == "8":
            fornecedores = gestao.listar_fornecedores()
            if not fornecedores:
                print("Nenhum fornecedor cadastrado.")
            for id_f, nome, celular, produto in fornecedores:
                print(f"ID {id_f} | {nome} | {celular} | {produto}")

        elif opcao == "9":
            lista_vendas = vendas.listar_vendas()
            if not lista_vendas:
                print("Nenhuma venda registrada.")
            for id_v, produto, quantidade, preco_unit, valor_total, data in lista_vendas:
                print(f"ID {id_v} | {quantidade}x {produto} | R${valor_total:.2f} | {data}")

        elif opcao == "10":
            total = gestao.valor_total_estoque()
            print(f"Valor total do estoque: R${total:.2f}")

        elif opcao == "11":
            produtos = gestao.lista_estoque()
            if not produtos:
                print("Estoque vazio, nao ha produtos para atualizar.")
            else:
                cursor = gestao.conn.cursor()
                cursor.execute("SELECT id, produto, quantidade, preco FROM estoque")
                for id_p, produto, quantidade, preco in cursor.fetchall():
                    print(f"ID {id_p} | {produto} | quantidade: {quantidade} | preco: R${preco:.2f}")

                produto_id = pedir_inteiro("\nDigite o ID do produto: ")
                quantidade = pedir_inteiro("Quantidade a adicionar: ")
                resposta = input("Atualizar o preco tambem? (s/n): ").strip().lower()

                if resposta == "s":
                    preco = pedir_float("Novo preco (ex: 10.50): ")
                else:
                    preco = None

                gestao.adicionar_produto_por_id(produto_id, quantidade, preco)

                produto_atualizado = gestao.consultar_produto_por_id(produto_id)
                if produto_atualizado:
                    nome_produto = produto_atualizado[1]
                    verificar_e_alertar(gestao, nome_produto, email_usuario)

        elif opcao == "0":
            print("\nGerando grafico atualizado do estoque...")
            gestao.gerar_grafico_estoque("relatorio_estoque.xlsx")
            print("Encerrando o programa.")
            break

        else:
            print("Opcao invalida, tente novamente.")