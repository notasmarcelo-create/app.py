import sqlite3
import pandas as pd
from datetime import datetime
import hashlib
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rebanho.db")

PRODUTORES_PADRAO = ["Michele", "Marcelo"]

USUARIOS_PADRAO = [
    ("michele", "michele123", "Michele", "Michele"),
    ("marcelo", "marcelo123", "Marcelo", "Marcelo"),
    ("admin", "admin123", "Administrador", "Todos os Produtores")
]

CATEGORIAS_PADRAO = [
    ("TERNEIROS 0-12", "0 a 12 meses", "Macho"),
    ("TERNEIRAS 0-12", "0 a 12 meses", "Fêmea"),
    ("NOVILHOS 0-12", "0 a 12 meses", "Macho"),
    ("NOVILHAS 0-12", "0 a 12 meses", "Fêmea"),
    ("MACHO 13-24", "13 a 24 meses", "Macho"),
    ("NOVILHOS 13-24", "13 a 24 meses", "Macho"),
    ("NOVILHAS 13-24", "13 a 24 meses", "Fêmea"),
    ("NOVILHOS 25-36", "25 a 36 meses", "Macho"),
    ("NOVILHAS 25-36", "25 a 36 meses", "Fêmea"),
    ("VACAS 25-36", "25 a 36 meses", "Fêmea"),
    ("VACAS +36", "Acima de 36 meses", "Fêmea"),
    ("MACHO +36", "Acima de 36 meses", "Macho"),
    ("TOUROS", "Acima de 24 meses", "Macho"),
    ("DIVERSOS / OUTROS", "Todas", "Misto"),
]

TIPOS_LANCAMENTO_PADRAO = [
    "SAÍDA PARA ABATE",
    "SAIDA P/ CRIA / ENGORDA",
    "ENTRADA P/ CRIA / ENGORDA",
    "NASCIMENTO",
    "MORTE",
    "CONSUMO",
    "TRANSFERÊNCIA",
    "AJUSTE DE INVENTÁRIO",
    "OUTROS"
]

def hash_senha(senha):
    """Gera hash SHA-256 seguro da senha."""
    return hashlib.sha256(senha.encode("utf-8")).hexdigest()

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    # Tabela de Usuários para Login
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        nome_completo TEXT NOT NULL,
        produtor_associado TEXT NOT NULL DEFAULT 'Todos os Produtores',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Tabela de Produtores
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS produtores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT UNIQUE NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Tabela de Categorias
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS categorias (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT UNIQUE NOT NULL,
        faixa_etaria TEXT,
        sexo TEXT
    )
    """)

    # Tabela de Movimentações
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS movimentacoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        produtor TEXT NOT NULL DEFAULT 'Michele',
        tipo_lancamento TEXT NOT NULL,
        data TEXT NOT NULL,
        operacao TEXT NOT NULL,
        qtd_total INTEGER NOT NULL,
        gta TEXT,
        nfp TEXT,
        observacoes TEXT,
        resumo_categorias TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cursor.execute("PRAGMA table_info(movimentacoes)")
    colunas = [info[1] for info in cursor.fetchall()]
    if "produtor" not in colunas:
        cursor.execute("ALTER TABLE movimentacoes ADD COLUMN produtor TEXT NOT NULL DEFAULT 'Michele'")

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS detalhes_categoria (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        movimentacao_id INTEGER NOT NULL,
        categoria TEXT NOT NULL,
        quantidade INTEGER NOT NULL,
        FOREIGN KEY (movimentacao_id) REFERENCES movimentacoes (id) ON DELETE CASCADE
    )
    """)

    # Popula usuários padrão se não existirem
    cursor.execute("SELECT COUNT(*) FROM usuarios")
    if cursor.fetchone()[0] == 0:
        for u, p, nome, prod_assoc in USUARIOS_PADRAO:
            cursor.execute(
                "INSERT OR IGNORE INTO usuarios (username, password_hash, nome_completo, produtor_associado) VALUES (?, ?, ?, ?)",
                (u.lower(), hash_senha(p), nome, prod_assoc)
            )

    # Popula produtores padrão
    for prod in PRODUTORES_PADRAO:
        cursor.execute("INSERT OR IGNORE INTO produtores (nome) VALUES (?)", (prod,))

    # Popula categorias padrão se vazio
    cursor.execute("SELECT COUNT(*) FROM categorias")
    if cursor.fetchone()[0] == 0:
        for nome, faixa, sexo in CATEGORIAS_PADRAO:
            cursor.execute(
                "INSERT OR IGNORE INTO categorias (nome, faixa_etaria, sexo) VALUES (?, ?, ?)",
                (nome, faixa, sexo)
            )

    conn.commit()
    conn.close()

# Funções de Autenticação / Usuários
def autenticar_usuario(username, senha):
    """Verifica credenciais de login. Retorna dicionário do usuário se válido, ou None."""
    conn = get_connection()
    cursor = conn.cursor()
    pwd_hash = hash_senha(senha)
    cursor.execute("SELECT * FROM usuarios WHERE LOWER(username) = ? AND password_hash = ?", (username.lower().strip(), pwd_hash))
    row = cursor.fetchone()
    conn.close()
    if row:
        return dict(row)
    return None

def alterar_senha_usuario(username, nova_senha):
    """Altera a senha de um usuário."""
    conn = get_connection()
    cursor = conn.cursor()
    pwd_hash = hash_senha(nova_senha)
    cursor.execute("UPDATE usuarios SET password_hash = ? WHERE LOWER(username) = ?", (pwd_hash, username.lower().strip()))
    conn.commit()
    conn.close()

def get_usuarios_cadastrados():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, nome_completo, produtor_associado FROM usuarios ORDER BY id ASC")
    users = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return users

def add_novo_usuario(username, senha, nome_completo, produtor_associado="Todos os Produtores"):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            "INSERT INTO usuarios (username, password_hash, nome_completo, produtor_associado) VALUES (?, ?, ?, ?)",
            (username.lower().strip(), hash_senha(senha), nome_completo.strip(), produtor_associado)
        )
        conn.commit()
        sucesso = True
    except sqlite3.IntegrityError:
        sucesso = False
    conn.close()
    return sucesso

def get_produtores():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT nome FROM produtores ORDER BY id ASC")
    prods = [r["nome"] for r in cursor.fetchall()]
    conn.close()
    if not prods:
        return PRODUTORES_PADRAO
    return prods

def add_produtor(nome):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO produtores (nome) VALUES (?)", (nome.strip(),))
        conn.commit()
        sucesso = True
    except sqlite3.IntegrityError:
        sucesso = False
    conn.close()
    return sucesso

def delete_produtor(nome):
    if nome in PRODUTORES_PADRAO and len(get_produtores()) <= 2:
        return False
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM produtores WHERE nome = ?", (nome,))
    conn.commit()
    conn.close()
    return True

def formatar_data_para_banco(data_str_ou_obj):
    if isinstance(data_str_ou_obj, str):
        data_str = data_str_ou_obj.strip()
        if "/" in data_str:
            partes = data_str.split("/")
            if len(partes) == 3:
                d, m, y = partes
                return f"{y.zfill(4)}-{m.zfill(2)}-{d.zfill(2)}"
        return data_str
    elif hasattr(data_str_ou_obj, "strftime"):
        return data_str_ou_obj.strftime("%Y-%m-%d")
    return str(data_str_ou_obj)

def formatar_data_para_exibicao(data_iso):
    if not data_iso:
        return ""
    if "-" in str(data_iso):
        partes = str(data_iso).split("-")
        if len(partes) == 3:
            y, m, d = partes
            return f"{d}/{m}/{y}"
    return str(data_iso)

def gerar_resumo_categorias(detalhes):
    if not detalhes:
        return ""
    if len(detalhes) == 1 and detalhes[0][1] == 0:
        return detalhes[0][0]
    
    partes = []
    for cat, qtd in detalhes:
        if qtd > 0:
            partes.append(f"{qtd} - {cat}")
        else:
            partes.append(f"{cat}")
    return " | ".join(partes)

def add_movimentacao(tipo_lancamento, data, operacao, qtd_total, produtor="Michele", gta="", nfp="", detalhes=None, observacoes=""):
    conn = get_connection()
    cursor = conn.cursor()

    data_iso = formatar_data_para_banco(data)
    operacao_limpa = operacao.upper().strip()
    produtor_limpo = produtor.strip() if produtor else "Michele"

    lista_detalhes = []
    if detalhes:
        for item in detalhes:
            if isinstance(item, dict):
                cat = item.get("categoria", "").strip()
                qtd = int(item.get("quantidade", 0))
                if cat:
                    lista_detalhes.append((cat, qtd))
            elif isinstance(item, (list, tuple)):
                cat = str(item[0]).strip()
                qtd = int(item[1]) if len(item) > 1 else 0
                if cat:
                    lista_detalhes.append((cat, qtd))

    if not lista_detalhes:
        lista_detalhes.append(("DIVERSOS / OUTROS", qtd_total))

    resumo_cat = gerar_resumo_categorias(lista_detalhes)

    cursor.execute("""
    INSERT INTO movimentacoes (produtor, tipo_lancamento, data, operacao, qtd_total, gta, nfp, observacoes, resumo_categorias)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (produtor_limpo, tipo_lancamento.strip(), data_iso, operacao_limpa, int(qtd_total), gta.strip() if gta else "", nfp.strip() if nfp else "", observacoes.strip() if observacoes else "", resumo_cat))

    mov_id = cursor.lastrowid

    for cat, qtd in lista_detalhes:
        cursor.execute("""
        INSERT INTO detalhes_categoria (movimentacao_id, categoria, quantidade)
        VALUES (?, ?, ?)
        """, (mov_id, cat, qtd))

    conn.commit()
    conn.close()
    return mov_id

def update_movimentacao(mov_id, tipo_lancamento, data, operacao, qtd_total, produtor="Michele", gta="", nfp="", detalhes=None, observacoes=""):
    conn = get_connection()
    cursor = conn.cursor()

    data_iso = formatar_data_para_banco(data)
    operacao_limpa = operacao.upper().strip()
    produtor_limpo = produtor.strip() if produtor else "Michele"

    lista_detalhes = []
    if detalhes:
        for item in detalhes:
            if isinstance(item, dict):
                cat = item.get("categoria", "").strip()
                qtd = int(item.get("quantidade", 0))
                if cat:
                    lista_detalhes.append((cat, qtd))
            elif isinstance(item, (list, tuple)):
                cat = str(item[0]).strip()
                qtd = int(item[1]) if len(item) > 1 else 0
                if cat:
                    lista_detalhes.append((cat, qtd))

    if not lista_detalhes:
        lista_detalhes.append(("DIVERSOS / OUTROS", qtd_total))

    resumo_cat = gerar_resumo_categorias(lista_detalhes)

    cursor.execute("""
    UPDATE movimentacoes
    SET produtor = ?, tipo_lancamento = ?, data = ?, operacao = ?, qtd_total = ?, gta = ?, nfp = ?, observacoes = ?, resumo_categorias = ?
    WHERE id = ?
    """, (produtor_limpo, tipo_lancamento.strip(), data_iso, operacao_limpa, int(qtd_total), gta.strip() if gta else "", nfp.strip() if nfp else "", observacoes.strip() if observacoes else "", resumo_cat, mov_id))

    cursor.execute("DELETE FROM detalhes_categoria WHERE movimentacao_id = ?", (mov_id,))

    for cat, qtd in lista_detalhes:
        cursor.execute("""
        INSERT INTO detalhes_categoria (movimentacao_id, categoria, quantidade)
        VALUES (?, ?, ?)
        """, (mov_id, cat, qtd))

    conn.commit()
    conn.close()

def delete_movimentacao(mov_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM movimentacoes WHERE id = ?", (mov_id,))
    conn.commit()
    conn.close()

def get_movimentacao(mov_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM movimentacoes WHERE id = ?", (mov_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return None

    mov = dict(row)
    cursor.execute("SELECT categoria, quantidade FROM detalhes_categoria WHERE movimentacao_id = ?", (mov_id,))
    mov["detalhes"] = [{"categoria": r["categoria"], "quantidade": r["quantidade"]} for r in cursor.fetchall()]
    conn.close()
    return mov

def get_todas_movimentacoes(filtros=None):
    conn = get_connection()
    query = """
    SELECT 
        m.id,
        m.produtor AS "PRODUTOR",
        m.tipo_lancamento AS "LANÇAMENTO",
        m.data AS "DATA_ISO",
        m.operacao AS "OPERAÇÃO",
        m.qtd_total AS "QTD. TOTAL",
        m.gta AS "GTA",
        m.nfp AS "NFP",
        m.resumo_categorias AS "ANIMAIS / MESES",
        m.observacoes AS "OBSERVAÇÕES"
    FROM movimentacoes m
    WHERE 1=1
    """
    params = []

    if filtros:
        if filtros.get("produtor") and filtros["produtor"] not in ("TODOS", "Todos os Produtores", "TODOS OS PRODUTORES"):
            query += " AND m.produtor = ?"
            params.append(filtros["produtor"])
        if filtros.get("data_inicio"):
            data_ini_iso = formatar_data_para_banco(filtros["data_inicio"])
            query += " AND m.data >= ?"
            params.append(data_ini_iso)
        if filtros.get("data_fim"):
            data_fim_iso = formatar_data_para_banco(filtros["data_fim"])
            query += " AND m.data <= ?"
            params.append(data_fim_iso)
        if filtros.get("tipo_lancamento") and filtros["tipo_lancamento"] != "TODOS":
            query += " AND m.tipo_lancamento = ?"
            params.append(filtros["tipo_lancamento"])
        if filtros.get("operacao") and filtros["operacao"] != "TODOS":
            query += " AND m.operacao = ?"
            params.append(filtros["operacao"])
        if filtros.get("busca_texto"):
            termo = f"%{filtros['busca_texto'].strip()}%"
            query += " AND (m.gta LIKE ? OR m.nfp LIKE ? OR m.resumo_categorias LIKE ? OR m.observacoes LIKE ? OR m.produtor LIKE ?)"
            params.extend([termo, termo, termo, termo, termo])
        if filtros.get("categoria") and filtros["categoria"] != "TODAS":
            query += " AND m.id IN (SELECT movimentacao_id FROM detalhes_categoria WHERE categoria = ?)"
            params.append(filtros["categoria"])

    query += " ORDER BY m.data DESC, m.id DESC"

    df = pd.read_sql_query(query, conn, params=params)
    conn.close()

    if not df.empty:
        df["DATA"] = df["DATA_ISO"].apply(formatar_data_para_exibicao)
    else:
        df["DATA"] = []

    return df

def get_resumo_totais(produtor=None):
    conn = get_connection()
    cursor = conn.cursor()

    query = """
    SELECT 
        COALESCE(SUM(CASE WHEN operacao = 'CRÉDITO' THEN qtd_total ELSE 0 END), 0) AS total_entradas,
        COALESCE(SUM(CASE WHEN operacao = 'DÉBITO' THEN qtd_total ELSE 0 END), 0) AS total_saidas,
        COUNT(*) AS total_registros
    FROM movimentacoes
    WHERE 1=1
    """
    params = []
    if produtor and produtor not in ("TODOS", "Todos os Produtores", "TODOS OS PRODUTORES"):
        query += " AND produtor = ?"
        params.append(produtor)

    cursor.execute(query, params)
    row = cursor.fetchone()
    conn.close()

    entradas = row["total_entradas"]
    saidas = row["total_saidas"]
    saldo_atual = entradas - saidas
    total_registros = row["total_registros"]

    return {
        "saldo_atual": saldo_atual,
        "total_entradas": entradas,
        "total_saidas": saidas,
        "total_registros": total_registros
    }

def get_saldo_por_categoria(produtor=None):
    conn = get_connection()
    cursor = conn.cursor()

    query = """
    SELECT 
        dc.categoria,
        SUM(CASE WHEN m.operacao = 'CRÉDITO' THEN dc.quantidade ELSE 0 END) AS entradas,
        SUM(CASE WHEN m.operacao = 'DÉBITO' THEN dc.quantidade ELSE 0 END) AS saidas,
        SUM(CASE WHEN m.operacao = 'CRÉDITO' THEN dc.quantidade ELSE -dc.quantidade END) AS saldo
    FROM detalhes_categoria dc
    JOIN movimentacoes m ON m.id = dc.movimentacao_id
    WHERE 1=1
    """
    params = []
    if produtor and produtor not in ("TODOS", "Todos os Produtores", "TODOS OS PRODUTORES"):
        query += " AND m.produtor = ?"
        params.append(produtor)

    query += """
    GROUP BY dc.categoria
    ORDER BY saldo DESC, dc.categoria ASC
    """

    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()

    dados = []
    for r in rows:
        dados.append({
            "Categoria": r["categoria"],
            "Entradas (Crédito)": r["entradas"],
            "Saídas (Débito)": r["saidas"],
            "Saldo Atual": r["saldo"]
        })

    return pd.DataFrame(dados)

def get_evolucao_mensal(produtor=None):
    conn = get_connection()
    query = """
    SELECT 
        substr(data, 1, 7) AS mes_ano,
        SUM(CASE WHEN operacao = 'CRÉDITO' THEN qtd_total ELSE 0 END) AS entradas,
        SUM(CASE WHEN operacao = 'DÉBITO' THEN qtd_total ELSE 0 END) AS saidas
    FROM movimentacoes
    WHERE 1=1
    """
    params = []
    if produtor and produtor not in ("TODOS", "Todos os Produtores", "TODOS OS PRODUTORES"):
        query += " AND produtor = ?"
        params.append(produtor)

    query += """
    GROUP BY mes_ano
    ORDER BY mes_ano ASC
    """
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()

    if not df.empty:
        df["saldo_mes"] = df["entradas"] - df["saidas"]
        df["saldo_acumulado"] = df["saldo_mes"].cumsum()
    return df

def get_resumo_por_produtores():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT 
        produtor,
        COALESCE(SUM(CASE WHEN operacao = 'CRÉDITO' THEN qtd_total ELSE 0 END), 0) AS entradas,
        COALESCE(SUM(CASE WHEN operacao = 'DÉBITO' THEN qtd_total ELSE 0 END), 0) AS saidas,
        COUNT(*) AS total_lancamentos
    FROM movimentacoes
    GROUP BY produtor
    ORDER BY produtor ASC
    """)
    rows = cursor.fetchall()
    conn.close()

    dados = []
    for r in rows:
        dados.append({
            "Produtor": r["produtor"],
            "Entradas": r["entradas"],
            "Saídas": r["saidas"],
            "Saldo Atual": r["entradas"] - r["saidas"],
            "Lançamentos": r["total_lancamentos"]
        })
    return pd.DataFrame(dados)

def get_categorias_cadastradas():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT nome FROM categorias ORDER BY nome ASC")
    cats = [r["nome"] for r in cursor.fetchall()]
    conn.close()
    return cats

def add_categoria_personalizada(nome, faixa_etaria="", sexo="Misto"):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO categorias (nome, faixa_etaria, sexo) VALUES (?, ?, ?)", (nome.strip().upper(), faixa_etaria, sexo))
        conn.commit()
        sucesso = True
    except sqlite3.IntegrityError:
        sucesso = False
    conn.close()
    return sucesso

def delete_categoria_personalizada(nome):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM categorias WHERE nome = ?", (nome,))
    conn.commit()
    conn.close()
