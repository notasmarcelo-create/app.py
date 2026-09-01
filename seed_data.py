from database import init_db, add_movimentacao, get_connection

DADOS_INICIAIS = [
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAÍDA PARA ABATE",
        "data": "08/12/2025",
        "operacao": "DÉBITO",
        "qtd_total": 26,
        "gta": "AD-616710",
        "nfp": "64117201",
        "detalhes": [("VACAS +36", 26)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAÍDA PARA ABATE",
        "data": "04/11/2025",
        "operacao": "DÉBITO",
        "qtd_total": 24,
        "gta": "AD-420267",
        "nfp": "63514672",
        "detalhes": [("VACAS +36", 24)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAÍDA PARA ABATE",
        "data": "03/11/2025",
        "operacao": "DÉBITO",
        "qtd_total": 7,
        "gta": "AD-415183",
        "nfp": "63491499",
        "detalhes": [("VACAS +36", 7)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAIDA P/ CRIA / ENGORDA",
        "data": "01/07/2025",
        "operacao": "DÉBITO",
        "qtd_total": 166,
        "gta": "AC-731844",
        "nfp": "61415793",
        "detalhes": [("TERNEIROS 0-12", 166)],
        "observacoes": "Data original na planilha: 01/07/2055 (ajustada para 2025)"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAÍDA PARA ABATE",
        "data": "10/03/2025",
        "operacao": "DÉBITO",
        "qtd_total": 10,
        "gta": "AC-167972",
        "nfp": "59474395",
        "detalhes": [("VACAS +36", 10)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAÍDA PARA ABATE",
        "data": "13/01/2025",
        "operacao": "DÉBITO",
        "qtd_total": 9,
        "gta": "AB-865384",
        "nfp": "58600527",
        "detalhes": [("VACAS +36", 9)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAÍDA PARA ABATE",
        "data": "04/12/2024",
        "operacao": "DÉBITO",
        "qtd_total": 8,
        "gta": "AB-668943",
        "nfp": "58059579",
        "detalhes": [("MACHO 13-24", 8)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAIDA P/ CRIA / ENGORDA",
        "data": "03/12/2024",
        "operacao": "DÉBITO",
        "qtd_total": 47,
        "gta": "AB-665752",
        "nfp": "58045787",
        "detalhes": [("NOVILHAS 13-24", 47)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "ENTRADA P/ CRIA / ENGORDA",
        "data": "31/10/2024",
        "operacao": "CRÉDITO",
        "qtd_total": 368,
        "gta": "AB-488470",
        "nfp": "57513259",
        "detalhes": [
            ("VACAS 25-36", 52),
            ("NOVILHAS 0-12", 141),
            ("NOVILHAS 13-24", 175)
        ],
        "observacoes": "Lote composto Michele (52 Vacas 25-36 + 141 Novilhas 0-12 + 175 Novilhas 13-24)"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAIDA P/ CRIA / ENGORDA",
        "data": "31/10/2024",
        "operacao": "DÉBITO",
        "qtd_total": 186,
        "gta": "AB-486984",
        "nfp": "57497242",
        "detalhes": [
            ("NOVILHAS 0-12", 141),
            ("NOVILHAS 13-24", 45)
        ],
        "observacoes": "Lote composto Michele (141 Novilhas 0-12 + 45 Novilhas 13-24)"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAIDA P/ CRIA / ENGORDA",
        "data": "28/10/2024",
        "operacao": "DÉBITO",
        "qtd_total": 43,
        "gta": "AB-467452",
        "nfp": "57428881",
        "detalhes": [("NOVILHOS 0-12", 43)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "ENTRADA P/ CRIA / ENGORDA",
        "data": "19/08/2024",
        "operacao": "CRÉDITO",
        "qtd_total": 548,
        "gta": "AB-060947",
        "nfp": "",
        "detalhes": [("DIVERSOS / OUTROS", 548)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAIDA P/ CRIA / ENGORDA",
        "data": "15/08/2024",
        "operacao": "DÉBITO",
        "qtd_total": 148,
        "gta": "AB-041368",
        "nfp": "56295372",
        "detalhes": [("NOVILHAS 0-12", 148)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAIDA P/ CRIA / ENGORDA",
        "data": "16/06/2024",
        "operacao": "DÉBITO",
        "qtd_total": 150,
        "gta": "AA-800204",
        "nfp": "",
        "detalhes": [("DIVERSOS / OUTROS", 150)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAIDA P/ CRIA / ENGORDA",
        "data": "19/06/2024",
        "operacao": "DÉBITO",
        "qtd_total": 64,
        "gta": "AA-800049",
        "nfp": "",
        "detalhes": [("DIVERSOS / OUTROS", 64)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAIDA P/ CRIA / ENGORDA",
        "data": "02/05/2024",
        "operacao": "DÉBITO",
        "qtd_total": 202,
        "gta": "AA-696763",
        "nfp": "",
        "detalhes": [("DIVERSOS / OUTROS", 202)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAÍDA PARA ABATE",
        "data": "09/04/2024",
        "operacao": "DÉBITO",
        "qtd_total": 26,
        "gta": "AA-582749",
        "nfp": "",
        "detalhes": [("DIVERSOS / OUTROS", 26)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "SAIDA P/ CRIA / ENGORDA",
        "data": "09/08/2023",
        "operacao": "DÉBITO",
        "qtd_total": 54,
        "gta": "Z-404346",
        "nfp": "",
        "detalhes": [("DIVERSOS / OUTROS", 54)],
        "observacoes": "Lançamento inicial Michele"
    },
    {
        "produtor": "Michele",
        "tipo_lancamento": "ENTRADA P/ CRIA / ENGORDA",
        "data": "30/06/2023",
        "operacao": "CRÉDITO",
        "qtd_total": 600,
        "gta": "Z-251461",
        "nfp": "",
        "detalhes": [("DIVERSOS / OUTROS", 600)],
        "observacoes": "Lançamento inicial Michele"
    }
]

def popular_banco(forcar=False):
    init_db()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM movimentacoes")
    qtd_existente = cursor.fetchone()[0]
    conn.close()

    if qtd_existente > 0 and not forcar:
        return False

    if forcar and qtd_existente > 0:
        conn = get_connection()
        conn.execute("DELETE FROM detalhes_categoria")
        conn.execute("DELETE FROM movimentacoes")
        conn.commit()
        conn.close()

    for item in DADOS_INICIAIS:
        add_movimentacao(
            tipo_lancamento=item["tipo_lancamento"],
            data=item["data"],
            operacao=item["operacao"],
            qtd_total=item["qtd_total"],
            produtor=item["produtor"],
            gta=item["gta"],
            nfp=item["nfp"],
            detalhes=item["detalhes"],
            observacoes=item["observacoes"]
        )

    return True

if __name__ == "__main__":
    popular_banco(forcar=True)
