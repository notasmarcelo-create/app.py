import streamlit as st
import pandas as pd
import datetime
from database import (
    init_db, get_todas_movimentacoes, get_resumo_totais, 
    get_saldo_por_categoria, get_evolucao_mensal, get_resumo_por_produtores,
    add_movimentacao, update_movimentacao, delete_movimentacao, get_movimentacao,
    get_categorias_cadastradas, add_categoria_personalizada, delete_categoria_personalizada,
    get_produtores, add_produtor, delete_produtor,
    autenticar_usuario, alterar_senha_usuario, get_usuarios_cadastrados, add_novo_usuario,
    TIPOS_LANCAMENTO_PADRAO, formatar_data_para_exibicao
)
from export_utils import gerar_relatorio_excel_bytes, importar_planilha_excel
from seed_data import popular_banco

# Configuração da Página
st.set_page_config(
    page_title="Controle de Rebanho - Marcelo & Michele",
    page_icon="🐄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inicializa o Banco
init_db()

# Inicializa Estado de Autenticação
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario_atual" not in st.session_state:
    st.session_state.usuario_atual = None

# Estilos CSS Personalizados
st.markdown("""
<style>
    .main-header {
        font-size: 26px;
        font-weight: 700;
        color: #1b4332;
        margin-bottom: 2px;
    }
    .sub-header {
        font-size: 14px;
        color: #555;
        margin-bottom: 15px;
    }
    .login-container {
        max-width: 420px;
        margin: 40px auto;
        padding: 30px;
        background: #ffffff;
        border-radius: 12px;
        box-shadow: 0 6px 20px rgba(0,0,0,0.08);
        border-top: 6px solid #2d6a4f;
    }
    .producer-badge {
        display: inline-block;
        padding: 5px 14px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 600;
        background-color: #e8f5e9;
        color: #2e7d32;
        border: 1px solid #c8e6c9;
        margin-bottom: 10px;
    }
    .metric-card {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        border-radius: 10px;
        padding: 16px;
        border-left: 5px solid #2d6a4f;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }
    .metric-title {
        font-size: 12px;
        font-weight: 600;
        color: #495057;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .metric-value {
        font-size: 26px;
        font-weight: 700;
        color: #1b4332;
        margin-top: 4px;
    }
    .prod-card {
        background: #f1f8f5;
        border: 1px solid #d8ebd4;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# TELA DE LOGIN (Se não estiver autenticado)
# ==========================================
if not st.session_state.autenticado:
    col_l1, col_l2, col_l3 = st.columns([1, 1.4, 1])
    with col_l2:
        st.write("")
        st.write("")
        st.markdown("""
        <div style="text-align: center; margin-bottom: 20px;">
            <div style="font-size: 48px;">🐄</div>
            <h2 style="color: #1b4332; margin-top: 5px; font-weight: 700;">Gestão de Rebanho</h2>
            <p style="color: #666; font-size: 14px;">Lançamentos de Gado • Marcelo & Michele</p>
        </div>
        """, unsafe_allow_html=True)

        with st.form("form_login"):
            st.markdown("#### 🔒 Acesso ao Sistema")
            username_input = st.text_input("Usuário", placeholder="michele, marcelo ou admin")
            senha_input = st.text_input("Senha", type="password", placeholder="Digite sua senha")
            
            btn_entrar = st.form_submit_button("🚀 Entrar no Sistema", use_container_width=True, type="primary")

            if btn_entrar:
                if username_input.strip() and senha_input.strip():
                    usuario_valido = autenticar_usuario(username_input, senha_input)
                    if usuario_valido:
                        st.session_state.autenticado = True
                        st.session_state.usuario_atual = usuario_valido
                        # Define produtor inicial conforme o usuário
                        if usuario_valido["produtor_associado"] in ["Michele", "Marcelo"]:
                            st.session_state.produtor_ativo = usuario_valido["produtor_associado"]
                        else:
                            st.session_state.produtor_ativo = "Todos os Produtores"
                        st.success(f"Bem-vindo(a), {usuario_valido['nome_completo']}!")
                        st.rerun()
                    else:
                        st.error("❌ Usuário ou senha incorretos. Tente novamente.")
                else:
                    st.warning("Preencha usuário e senha.")

        # Dica com os logins padrão
        with st.expander("ℹ️ Acessos Padrão Configurados"):
            st.markdown("""
            | Usuário | Senha Padrão | Produtor Associado |
            | :--- | :--- | :--- |
            | **michele** | `michele123` | Michele |
            | **marcelo** | `marcelo123` | Marcelo |
            | **admin** | `admin123` | Todos os Produtores |
            
            *(Você poderá alterar a senha depois de entrar no sistema)*
            """)
    st.stop()

# ==========================================
# SISTEMA PRINCIPAL (APÓS LOGIN)
# ==========================================

usuario_logado = st.session_state.usuario_atual
lista_produtores = get_produtores()
if not lista_produtores:
    lista_produtores = ["Michele", "Marcelo"]

# Sidebar - Usuário Logado e Seleção Global de Produtor
with st.sidebar:
    st.markdown(f"### 👤 Olá, {usuario_logado['nome_completo']}!")
    st.caption(f"Usuário: `{usuario_logado['username']}`")
    
    if st.button("🚪 Sair / Desconectar", use_container_width=True):
        st.session_state.autenticado = False
        st.session_state.usuario_atual = None
        st.rerun()

    st.divider()

    st.markdown("#### 🚜 Visualização de Produtor")
    opcoes_produtor_global = ["Todos os Produtores"] + lista_produtores
    
    if "produtor_ativo" not in st.session_state:
        st.session_state.produtor_ativo = "Todos os Produtores"

    produtor_selecionado = st.selectbox(
        "Visualizar dados de:",
        opcoes_produtor_global,
        index=opcoes_produtor_global.index(st.session_state.produtor_ativo) if st.session_state.produtor_ativo in opcoes_produtor_global else 0,
        key="select_produtor_global"
    )
    st.session_state.produtor_ativo = produtor_selecionado

    st.markdown("---")
    with st.expander("🔑 Alterar Minha Senha"):
        with st.form("form_troca_senha_sidebar"):
            nova_pwd = st.text_input("Nova Senha", type="password")
            conf_pwd = st.text_input("Confirmar Nova Senha", type="password")
            btn_trocar = st.form_submit_button("Salvar Nova Senha", type="primary")
            if btn_trocar:
                if nova_pwd and nova_pwd == conf_pwd:
                    alterar_senha_usuario(usuario_logado["username"], nova_pwd)
                    st.success("Senha alterada com sucesso!")
                else:
                    st.error("As senhas não coincidem ou estão vazias.")

    st.caption("Sistema de Gestão Agropecuária v2.0 • Protegido com Login")

# Top Bar / Título Principal
col_head1, col_head2 = st.columns([3, 1])
with col_head1:
    st.markdown('<div class="main-header">🐄 Sistema de Gestão de Rebanho & Lançamentos</div>', unsafe_allow_html=True)
    if produtor_selecionado != "Todos os Produtores":
        st.markdown(f'<div class="producer-badge">👤 Filtrando por: <b>{produtor_selecionado}</b></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="producer-badge" style="background-color: #e3f2fd; color: #1565c0; border-color: #bbdefb;">👥 <b>Visão Consolidada (Michele e Marcelo)</b></div>', unsafe_allow_html=True)
with col_head2:
    excel_data_quick = gerar_relatorio_excel_bytes({"produtor": produtor_selecionado})
    st.download_button(
        label="📥 Baixar Excel",
        data=excel_data_quick,
        file_name=f"rebanho_{produtor_selecionado.lower().replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# Abas Principais
tab_dashboard, tab_novo, tab_extrato, tab_import_export, tab_ajustes = st.tabs([
    "📊 Painel Geral & Estoque",
    "➕ Novo Lançamento",
    "📋 Histórico & Extrato",
    "📤 Importar / Exportar",
    "⚙️ Usuários & Configurações"
])

# ==========================================
# ABA 1: PAINEL GERAL & ESTOQUE
# ==========================================
with tab_dashboard:
    resumo = get_resumo_totais(produtor_selecionado)
    df_saldo_cat = get_saldo_por_categoria(produtor_selecionado)
    df_evolucao = get_evolucao_mensal(produtor_selecionado)
    df_resumo_prods = get_resumo_por_produtores()

    # Cards de KPIs
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="metric-card" style="border-left-color: #2d6a4f;">
            <div class="metric-title">🐂 Saldo Atual do Rebanho</div>
            <div class="metric-value">{resumo['saldo_atual']:,} <span style="font-size: 13px; font-weight: normal; color: #666;">cabeças</span></div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="metric-card" style="border-left-color: #2b9348;">
            <div class="metric-title">📥 Total Entradas (Crédito)</div>
            <div class="metric-value" style="color: #2b9348;">+{resumo['total_entradas']:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="metric-card" style="border-left-color: #d90429;">
            <div class="metric-title">📤 Total Saídas (Débito)</div>
            <div class="metric-value" style="color: #d90429;">-{resumo['total_saidas']:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="metric-card" style="border-left-color: #0077b6;">
            <div class="metric-title">📄 Total Lançamentos</div>
            <div class="metric-value" style="color: #0077b6;">{resumo['total_registros']}</div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    
    if produtor_selecionado == "Todos os Produtores" and not df_resumo_prods.empty:
        st.markdown("#### 👥 Resumo Comparativo por Produtor")
        cols_prods = st.columns(len(df_resumo_prods))
        for idx, row_p in df_resumo_prods.iterrows():
            with cols_prods[idx % len(cols_prods)]:
                st.markdown(f"""
                <div class="prod-card">
                    <div style="font-size: 15px; font-weight: bold; color: #1b4332;">👤 {row_p['Produtor']}</div>
                    <div style="font-size: 20px; font-weight: bold; color: #2d6a4f; margin: 4px 0;">{row_p['Saldo Atual']:,} <span style="font-size: 12px; font-weight: normal; color: #555;">cab. no estoque</span></div>
                    <div style="font-size: 12px; color: #555;">Entradas: <b>+{row_p['Entradas']}</b> | Saídas: <b>-{row_p['Saídas']}</b> ({row_p['Lançamentos']} docs)</div>
                </div>
                """, unsafe_allow_html=True)

    st.divider()

    col_cat, col_graf = st.columns([1.1, 1.2])

    with col_cat:
        st.subheader(f"📦 Saldo por Categoria / Idade ({produtor_selecionado})")
        if not df_saldo_cat.empty:
            st.dataframe(
                df_saldo_cat,
                column_config={
                    "Categoria": st.column_config.TextColumn("Categoria / Idade", width="medium"),
                    "Entradas (Crédito)": st.column_config.NumberColumn("Entradas", format="%d"),
                    "Saídas (Débito)": st.column_config.NumberColumn("Saídas", format="%d"),
                    "Saldo Atual": st.column_config.NumberColumn("Saldo Atual", format="%d")
                },
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Nenhum saldo registrado para este produtor.")

    with col_graf:
        st.subheader(f"📈 Movimentações Mensais ({produtor_selecionado})")
        if not df_evolucao.empty:
            chart_data = df_evolucao.set_index("mes_ano")[["entradas", "saidas"]]
            chart_data.columns = ["Entradas (Crédito)", "Saídas (Débito)"]
            st.bar_chart(chart_data, color=["#2b9348", "#d90429"], height=320)
        else:
            st.info("Sem dados suficientes para gráficos.")

# ==========================================
# ABA 2: NOVO LANÇAMENTO
# ==========================================
with tab_novo:
    st.subheader("📝 Cadastrar Nova Movimentação")
    st.caption("Selecione o produtor rural (Marcelo ou Michele), dados do documento e detalhe os animais.")

    categorias_disponiveis = get_categorias_cadastradas()
    if not categorias_disponiveis:
        categorias_disponiveis = ["DIVERSOS / OUTROS", "VACAS +36", "NOVILHAS 13-24", "NOVILHAS 0-12", "TERNEIROS 0-12"]

    idx_prod_form = 0
    if produtor_selecionado in lista_produtores:
        idx_prod_form = lista_produtores.index(produtor_selecionado)

    if "num_linhas_categoria" not in st.session_state:
        st.session_state.num_linhas_categoria = 1

    with st.form("form_novo_lancamento", clear_on_submit=True):
        col_f0, col_f1, col_f2, col_f3 = st.columns([1.2, 1.5, 1, 1.2])
        with col_f0:
            prod_lancamento = st.selectbox("Produtor Rural *", lista_produtores, index=idx_prod_form)
        with col_f1:
            tipo_lanc = st.selectbox("Tipo de Lançamento *", TIPOS_LANCAMENTO_PADRAO)
        with col_f2:
            data_mov = st.date_input("Data *", value=datetime.date.today(), format="DD/MM/YYYY")
        with col_f3:
            operacao = st.selectbox(
                "Operação *",
                ["DÉBITO (Saída/Abate/Venda)", "CRÉDITO (Entrada/Compra/Nascimento)"]
            )

        col_f4, col_f5 = st.columns(2)
        with col_f4:
            gta = st.text_input("Número da GTA (Guia de Trânsito Animal)", placeholder="Ex: AD-616710")
        with col_f5:
            nfp = st.text_input("Número da NFP (Nota Fiscal de Produtor)", placeholder="Ex: 64117201")

        st.markdown("##### 🐄 Detalhamento dos Animais")
        st.caption("Informe as categorias e quantidades deste lote:")

        detalhes_form = []
        qtd_somada = 0

        for i in range(st.session_state.num_linhas_categoria):
            c_cat, c_qtd = st.columns([3, 1])
            with c_cat:
                cat_selecionada = st.selectbox(
                    f"Categoria / Faixa Etária #{i+1}",
                    categorias_disponiveis,
                    key=f"novo_cat_{i}"
                )
            with c_qtd:
                qtd_categoria = st.number_input(
                    f"Quantidade #{i+1}",
                    min_value=0,
                    value=0,
                    step=1,
                    key=f"novo_qtd_{i}"
                )
            if cat_selecionada and qtd_categoria > 0:
                detalhes_form.append((cat_selecionada, int(qtd_categoria)))
                qtd_somada += int(qtd_categoria)

        obs = st.text_area("Observações Adicionais", placeholder="Informações complementares sobre este lote...")

        col_btn_sub, col_info_total = st.columns([1.2, 2])
        with col_info_total:
            if qtd_somada > 0:
                st.info(f"🔢 **Total:** {qtd_somada} cabeças para **{prod_lancamento}**")
            else:
                st.warning("⚠️ Informe a quantidade de pelo menos uma categoria.")

        with col_btn_sub:
            submitted = st.form_submit_button("💾 Salvar Lançamento", use_container_width=True, type="primary")

        if submitted:
            if qtd_somada <= 0:
                st.error("A quantidade total de animais deve ser maior que zero!")
            else:
                op_banco = "CRÉDITO" if "CRÉDITO" in operacao else "DÉBITO"
                add_movimentacao(
                    produtor=prod_lancamento,
                    tipo_lancamento=tipo_lanc,
                    data=data_mov,
                    operacao=op_banco,
                    qtd_total=qtd_somada,
                    gta=gta,
                    nfp=nfp,
                    detalhes=detalhes_form,
                    observacoes=obs
                )
                st.success(f"✅ Lançamento de {qtd_somada} cabeças salvo com sucesso para {prod_lancamento}!")
                st.rerun()

    col_add_rem1, col_add_rem2, _ = st.columns([1, 1, 2])
    with col_add_rem1:
        if st.button("➕ Adicionar Mais Uma Categoria no Lote"):
            st.session_state.num_linhas_categoria += 1
            st.rerun()
    with col_add_rem2:
        if st.session_state.num_linhas_categoria > 1:
            if st.button("➖ Remover Linha"):
                st.session_state.num_linhas_categoria -= 1
                st.rerun()

# ==========================================
# ABA 3: HISTÓRICO & EXTRATO
# ==========================================
with tab_extrato:
    st.subheader("📋 Histórico Completo de Lançamentos")
    st.caption("Consulte, filtre, edite ou exclua lançamentos de Marcelo e Michele.")

    with st.expander("🔍 Filtros de Busca e Pesquisa", expanded=True):
        f_col0, f_col1, f_col2, f_col3, f_col4 = st.columns(5)
        with f_col0:
            f_prod = st.selectbox(
                "Produtor", 
                ["TODOS"] + lista_produtores,
                index=0 if produtor_selecionado == "Todos os Produtores" else (lista_produtores.index(produtor_selecionado) + 1 if produtor_selecionado in lista_produtores else 0)
            )
        with f_col1:
            tipo_filtro = st.selectbox("Lançamento", ["TODOS"] + TIPOS_LANCAMENTO_PADRAO)
        with f_col2:
            op_filtro = st.selectbox("Operação", ["TODOS", "CRÉDITO", "DÉBITO"])
        with f_col3:
            todas_cats = ["TODAS"] + get_categorias_cadastradas()
            cat_filtro = st.selectbox("Categoria", todas_cats)
        with f_col4:
            termo_busca = st.text_input("GTA / NFP / Texto", placeholder="Ex: AD-616710")

        f_data1, f_data2, _ = st.columns([1, 1, 2])
        with f_data1:
            usar_filtro_data = st.checkbox("Filtrar por Período de Datas")
        with f_data2:
            if usar_filtro_data:
                d_ini = st.date_input("Data Inicial", value=datetime.date(2023, 1, 1), format="DD/MM/YYYY")
                d_fim = st.date_input("Data Final", value=datetime.date.today() + datetime.timedelta(days=365), format="DD/MM/YYYY")
            else:
                d_ini = None
                d_fim = None

    filtros_dict = {
        "produtor": f_prod,
        "tipo_lancamento": tipo_filtro,
        "operacao": op_filtro,
        "categoria": cat_filtro,
        "busca_texto": termo_busca,
        "data_inicio": d_ini if usar_filtro_data else None,
        "data_fim": d_fim if usar_filtro_data else None
    }

    df_tabela = get_todas_movimentacoes(filtros_dict)

    if not df_tabela.empty:
        col_info_tabela, col_btn_tabela = st.columns([3, 1])
        with col_info_tabela:
            st.write(f"Mostrando **{len(df_tabela)}** lançamentos encontrados.")
        with col_btn_tabela:
            excel_filtrado = gerar_relatorio_excel_bytes(filtros_dict)
            st.download_button(
                label="📥 Exportar Dados Filtrados",
                data=excel_filtrado,
                file_name="lancamentos_filtrados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        colunas_mostrar = ["id", "PRODUTOR", "LANÇAMENTO", "DATA", "OPERAÇÃO", "QTD. TOTAL", "GTA", "NFP", "ANIMAIS / MESES", "OBSERVAÇÕES"]
        st.dataframe(
            df_tabela[colunas_mostrar],
            column_config={
                "id": st.column_config.NumberColumn("ID", width="small"),
                "PRODUTOR": st.column_config.TextColumn("Produtor", width="small"),
                "LANÇAMENTO": st.column_config.TextColumn("Lançamento", width="medium"),
                "DATA": st.column_config.TextColumn("Data", width="small"),
                "OPERAÇÃO": st.column_config.TextColumn("Operação", width="small"),
                "QTD. TOTAL": st.column_config.NumberColumn("Qtd. Total", format="%d", width="small"),
                "GTA": st.column_config.TextColumn("GTA", width="small"),
                "NFP": st.column_config.TextColumn("NFP", width="small"),
                "ANIMAIS / MESES": st.column_config.TextColumn("Animais / Meses", width="large"),
                "OBSERVAÇÕES": st.column_config.TextColumn("Observações", width="medium"),
            },
            use_container_width=True,
            hide_index=True
        )

        st.divider()
        st.subheader("⚙️ Ações no Registro (Editar / Excluir)")
        
        col_sel_id, col_btn_edit, col_btn_del = st.columns([2, 1, 1])
        with col_sel_id:
            ids_disponiveis = df_tabela["id"].tolist()
            id_selecionado = st.selectbox("Selecione o ID do Lançamento:", ids_disponiveis)
        
        with col_btn_del:
            st.write("")
            st.write("")
            if st.button("🗑️ Excluir Registro", type="secondary", use_container_width=True):
                delete_movimentacao(id_selecionado)
                st.success(f"Lançamento #{id_selecionado} excluído!")
                st.rerun()

        with col_btn_edit:
            st.write("")
            st.write("")
            editar_modo = st.checkbox("✏️ Editar Registro")

        if editar_modo and id_selecionado:
            mov_atual = get_movimentacao(id_selecionado)
            if mov_atual:
                st.markdown(f"#### Editando Lançamento #{id_selecionado}")
                with st.form(f"form_editar_{id_selecionado}"):
                    ce0, ce1, ce2, ce3 = st.columns(4)
                    with ce0:
                        idx_p = lista_produtores.index(mov_atual.get("produtor", "Michele")) if mov_atual.get("produtor") in lista_produtores else 0
                        edit_prod = st.selectbox("Produtor", lista_produtores, index=idx_p)
                    with ce1:
                        try:
                            data_obj = datetime.datetime.strptime(mov_atual["data"], "%Y-%m-%d").date()
                        except Exception:
                            data_obj = datetime.date.today()
                        edit_data = st.date_input("Data", value=data_obj, format="DD/MM/YYYY")
                    with ce2:
                        idx_tipo = TIPOS_LANCAMENTO_PADRAO.index(mov_atual["tipo_lancamento"]) if mov_atual["tipo_lancamento"] in TIPOS_LANCAMENTO_PADRAO else 0
                        edit_tipo = st.selectbox("Lançamento", TIPOS_LANCAMENTO_PADRAO, index=idx_tipo)
                    with ce3:
                        idx_op = 0 if mov_atual["operacao"] == "CRÉDITO" else 1
                        edit_op = st.selectbox("Operação", ["CRÉDITO", "DÉBITO"], index=idx_op)

                    ce4, ce5, ce6 = st.columns(3)
                    with ce4:
                        edit_qtd = st.number_input("Qtd Total", value=int(mov_atual["qtd_total"]), min_value=1)
                    with ce5:
                        edit_gta = st.text_input("GTA", value=mov_atual["gta"] or "")
                    with ce6:
                        edit_nfp = st.text_input("NFP", value=mov_atual["nfp"] or "")

                    edit_obs = st.text_area("Observações", value=mov_atual["observacoes"] or "")
                    
                    detalhes_atuais = [(d["categoria"], d["quantidade"]) for d in mov_atual.get("detalhes", [])]

                    salvar_edicao = st.form_submit_button("💾 Salvar Alterações", type="primary")
                    if salvar_edicao:
                        update_movimentacao(
                            mov_id=id_selecionado,
                            produtor=edit_prod,
                            tipo_lancamento=edit_tipo,
                            data=edit_data,
                            operacao=edit_op,
                            qtd_total=edit_qtd,
                            gta=edit_gta,
                            nfp=edit_nfp,
                            detalhes=detalhes_atuais if detalhes_atuais else [("DIVERSOS / OUTROS", edit_qtd)],
                            observacoes=edit_obs
                        )
                        st.success("Lançamento atualizado com sucesso!")
                        st.rerun()

    else:
        st.info("Nenhum lançamento encontrado com os filtros selecionados.")

# ==========================================
# ABA 4: IMPORTAR / EXPORTAR
# ==========================================
with tab_import_export:
    st.subheader("📤 Exportar e Importar Planilhas")

    col_exp, col_imp = st.columns(2)

    with col_exp:
        st.markdown("### 📥 Exportar para Excel (.xlsx)")
        st.write("Baixe a planilha completa com abas de Lançamentos, Saldo por Categoria e Resumo por Produtor.")
        
        prod_exp = st.selectbox("Exportar dados de:", ["Todos os Produtores"] + lista_produtores, key="exp_prod_sel")
        excel_bytes = gerar_relatorio_excel_bytes({"produtor": prod_exp})
        st.download_button(
            label=f"📊 Baixar Planilha ({prod_exp})",
            data=excel_bytes,
            file_name=f"lancamentos_gado_{prod_exp.lower().replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

        st.write("")
        st.divider()
        st.markdown("### 🔄 Restaurar Dados Iniciais da Foto")
        st.write("Deseja restaurar os 19 lançamentos originais da Michele?")
        if st.button("⚠️ Restaurar Lançamentos Padrão da Foto", type="secondary"):
            popular_banco(forcar=True)
            st.success("Dados padrão da planilha restaurados com sucesso para Michele!")
            st.rerun()

    with col_imp:
        st.markdown("### 📂 Importar Planilha Excel / CSV")
        st.write("Envie um arquivo `.xlsx` ou `.csv` para carregar novos lançamentos.")
        
        prod_imp_padrao = st.selectbox("Atribuir lançamentos sem produtor para:", lista_produtores, key="imp_prod_sel")
        uploaded_file = st.file_uploader("Selecione o arquivo Excel ou CSV", type=["xlsx", "xls", "csv"])
        if uploaded_file is not None:
            if st.button("🚀 Processar e Importar Planilha", type="primary", use_container_width=True):
                qtd_imp, erros = importar_planilha_excel(uploaded_file, produtor_padrao=prod_imp_padrao)
                if qtd_imp > 0:
                    st.success(f"🎉 {qtd_imp} lançamentos importados com sucesso!")
                if erros:
                    st.warning(f"{len(erros)} avisos durante a importação:")
                    for err in erros[:5]:
                        st.write(f"- {err}")
                st.rerun()

# ==========================================
# ABA 5: USUÁRIOS & CONFIGURAÇÕES
# ==========================================
with tab_ajustes:
    st.subheader("⚙️ Usuários, Produtores e Categorias")

    col_sec0, col_sec1, col_sec2 = st.columns(3)

    # Gestão de Usuários
    with col_sec0:
        st.markdown("### 👥 Usuários do Sistema")
        users = get_usuarios_cadastrados()
        df_u = pd.DataFrame(users)[["username", "nome_completo", "produtor_associado"]]
        df_u.columns = ["Usuário", "Nome", "Produtor Padrão"]
        st.dataframe(df_u, use_container_width=True, hide_index=True)

        with st.form("form_novo_usuario_cad"):
            st.markdown("##### Cadastrar Novo Usuário")
            new_u = st.text_input("Usuário (Login)", placeholder="Ex: joao")
            new_p = st.text_input("Senha", type="password", placeholder="Ex: senha123")
            new_nome = st.text_input("Nome Completo", placeholder="Ex: João da Silva")
            new_assoc = st.selectbox("Produtor Associado", ["Todos os Produtores"] + lista_produtores)
            
            btn_add_u = st.form_submit_button("➕ Criar Usuário", type="primary")
            if btn_add_u:
                if new_u.strip() and new_p.strip() and new_nome.strip():
                    if add_novo_usuario(new_u, new_p, new_nome, new_assoc):
                        st.success(f"Usuário '{new_u}' criado com sucesso!")
                        st.rerun()
                    else:
                        st.error("Nome de usuário já existe.")
                else:
                    st.warning("Preencha todos os campos do usuário.")

    # Gestão de Produtores
    with col_sec1:
        st.markdown("### 🌾 Produtores Rurais")
        df_p = pd.DataFrame({"Produtor Rural": lista_produtores})
        st.dataframe(df_p, use_container_width=True, hide_index=True)

        with st.form("form_novo_produtor"):
            st.markdown("##### Cadastrar Novo Produtor")
            novo_prod_nome = st.text_input("Nome do Produtor", placeholder="Ex: Fazenda Santa Maria")
            btn_add_p = st.form_submit_button("➕ Adicionar Produtor", type="primary")
            if btn_add_p:
                if novo_prod_nome.strip():
                    if add_produtor(novo_prod_nome):
                        st.success(f"Produtor '{novo_prod_nome.strip()}' adicionado!")
                        st.rerun()
                    else:
                        st.error("Este produtor já existe.")
                else:
                    st.warning("Informe o nome do produtor.")

    # Gestão de Categorias
    with col_sec2:
        st.markdown("### 🏷️ Categorias de Animais")
        cats = get_categorias_cadastradas()
        df_cats = pd.DataFrame({"Categoria Cadastrada": cats})
        st.dataframe(df_cats, use_container_width=True, hide_index=True)

        with st.form("form_nova_categoria"):
            st.markdown("##### Cadastrar Nova Categoria")
            nova_cat = st.text_input("Nome da Categoria", placeholder="Ex: BEZERROS DESMAMADOS")
            faixa = st.text_input("Faixa Etária", placeholder="Ex: 8 a 12 meses")
            sexo = st.selectbox("Sexo", ["Misto", "Macho", "Fêmea"])
            
            btn_add_cat = st.form_submit_button("➕ Adicionar Categoria", type="primary")
            if btn_add_cat:
                if nova_cat.strip():
                    if add_categoria_personalizada(nova_cat, faixa, sexo):
                        st.success(f"Categoria '{nova_cat.upper()}' cadastrada!")
                        st.rerun()
                    else:
                        st.error("Esta categoria já existe.")
                else:
                    st.warning("Informe o nome da categoria.")

# Rodapé
st.markdown("---")
st.markdown("<div style='text-align: center; color: #888; font-size: 12px;'>Sistema de Gestão de Rebanho Bovino • Controle Seguro de Marcelo & Michele</div>", unsafe_allow_html=True)
