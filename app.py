from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
import os
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "pelog_secret_key")

DATABASE_URL = os.environ.get("DATABASE_URL")

if not DATABASE_URL:
    DATABASE_URL = "dbname=pelog user=postgres password=123456 host=localhost port=5432"


def conectar():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)


FUSO_BRASIL = ZoneInfo("America/Sao_Paulo")


def agora_brasil():
    return datetime.now(FUSO_BRASIL).replace(tzinfo=None)


USUARIOS = {
    "admin": {"senha": "admin123", "tipo": "admin"},
    "portaria": {"senha": "portaria123", "tipo": "portaria"},
    "pedro_cd": {"senha": "123", "tipo": "encarregado"},
    "jean_cd": {"senha": "123", "tipo": "encarregado"},
    "jalison_cross": {"senha": "123", "tipo": "encarregado"},
    "elaine_cross": {"senha": "123", "tipo": "encarregado"},
}


def setor_do_usuario(usuario):
    if usuario in ["jalison_cross", "elaine_cross"]:
        return "CROSS"

    if usuario in ["pedro_cd", "jean_cd"]:
        return "CD"

    return None


def criar_colunas():
    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS caminhoes (
            id SERIAL PRIMARY KEY
        );
    """)

    colunas = [
        "ADD COLUMN IF NOT EXISTS placa VARCHAR(20)",
        "ADD COLUMN IF NOT EXISTS motorista VARCHAR(100)",
        "ADD COLUMN IF NOT EXISTS cpf VARCHAR(20)",
        "ADD COLUMN IF NOT EXISTS empresa VARCHAR(100)",
        "ADD COLUMN IF NOT EXISTS tipo_material VARCHAR(100)",
        "ADD COLUMN IF NOT EXISTS nota_fiscal VARCHAR(100)",
        "ADD COLUMN IF NOT EXISTS setor_doca VARCHAR(20)",
        "ADD COLUMN IF NOT EXISTS horario TIMESTAMP",
        "ADD COLUMN IF NOT EXISTS doca VARCHAR(50)",
        "ADD COLUMN IF NOT EXISTS status VARCHAR(50)",
        "ADD COLUMN IF NOT EXISTS autorizado_por VARCHAR(100)",
        "ADD COLUMN IF NOT EXISTS horario_autorizacao TIMESTAMP",
        "ADD COLUMN IF NOT EXISTS inicio_doca TIMESTAMP",
        "ADD COLUMN IF NOT EXISTS fim_doca TIMESTAMP",
        "ADD COLUMN IF NOT EXISTS dts_observacao TEXT",
        "ADD COLUMN IF NOT EXISTS produtos_sku VARCHAR(255)",
        "ADD COLUMN IF NOT EXISTS notas VARCHAR(255)",
        "ADD COLUMN IF NOT EXISTS diferenca_os VARCHAR(255)",
        "ADD COLUMN IF NOT EXISTS quantidade_nfs INTEGER",
        "ADD COLUMN IF NOT EXISTS qtd_paletes_nf INTEGER",
        "ADD COLUMN IF NOT EXISTS qtd_paletes_conferido INTEGER",
        "ADD COLUMN IF NOT EXISTS peso_kg NUMERIC(10,2)",
        "ADD COLUMN IF NOT EXISTS diferenca_produtos TEXT",
        "ADD COLUMN IF NOT EXISTS equipe TEXT",
        "ADD COLUMN IF NOT EXISTS operacional_cadastrado_por VARCHAR(100)",
        "ADD COLUMN IF NOT EXISTS horario_operacional TIMESTAMP"
    ]

    for coluna in colunas:
        cur.execute(f"ALTER TABLE caminhoes {coluna};")

    conn.commit()
    cur.close()
    conn.close()


@app.before_request
def iniciar():
    criar_colunas()


@app.route("/")
def index():
    if "usuario" not in session:
        return redirect(url_for("login"))

    if session.get("tipo") == "admin":
        return redirect(url_for("admin"))

    if session.get("tipo") == "portaria":
        return redirect(url_for("portaria"))

    if session.get("tipo") == "encarregado":
        return redirect(url_for("encarregado"))

    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario")
        senha = request.form.get("senha")

        if usuario in USUARIOS and USUARIOS[usuario]["senha"] == senha:
            session["usuario"] = usuario
            session["tipo"] = USUARIOS[usuario]["tipo"]
            return redirect(url_for("index"))

        return render_template("login.html", erro="Usuário ou senha inválidos")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/portaria")
def portaria():
    if session.get("tipo") != "portaria":
        return redirect(url_for("login"))

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM caminhoes
        WHERE status IS NULL OR status != 'finalizado'
        ORDER BY id DESC;
    """)

    caminhoes = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        "portaria.html",
        caminhoes=caminhoes,
        usuario=session.get("usuario"),
        perfil=session.get("tipo")
    )


@app.route("/api/portaria")
def api_portaria():
    if session.get("tipo") != "portaria":
        return jsonify([])

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            id, 
            placa, 
            motorista, 
            empresa, 
            tipo_material, 
            nota_fiscal,
            setor_doca,
            status, 
            doca
        FROM caminhoes
        WHERE status IS NULL OR status != 'finalizado'
        ORDER BY id DESC;
    """)

    caminhoes = cur.fetchall()

    cur.close()
    conn.close()

    return jsonify(caminhoes)


@app.route("/registrar", methods=["POST"])
def registrar():
    if session.get("tipo") != "portaria":
        return redirect(url_for("login"))

    setor_doca = request.form.get("setor_doca")

    if setor_doca not in ["CROSS", "CD"]:
        return redirect(url_for("portaria"))

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO caminhoes
        (
            placa, 
            motorista, 
            cpf, 
            empresa, 
            tipo_material, 
            nota_fiscal, 
            setor_doca,
            horario, 
            status
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'aguardando');
    """, (
        request.form.get("placa"),
        request.form.get("motorista"),
        request.form.get("cpf"),
        request.form.get("empresa"),
        request.form.get("tipo_material"),
        request.form.get("nota_fiscal"),
        setor_doca,
        agora_brasil()
    ))

    conn.commit()
    cur.close()
    conn.close()

    return redirect(url_for("portaria"))


@app.route("/encarregado")
def encarregado():
    if session.get("tipo") != "encarregado":
        return redirect(url_for("login"))

    usuario = session.get("usuario")
    setor = setor_do_usuario(usuario)

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM caminhoes
        WHERE setor_doca = %s
          AND (status IS NULL OR status != 'finalizado')
        ORDER BY id DESC;
    """, (setor,))

    caminhoes = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        "encarregado.html",
        caminhoes=caminhoes,
        usuario=usuario,
        perfil=session.get("tipo"),
        setor=setor
    )


@app.route("/autorizar/<int:id>", methods=["POST"])
def autorizar(id):
    if session.get("tipo") != "encarregado":
        return redirect(url_for("login"))

    usuario = session.get("usuario")
    setor = setor_do_usuario(usuario)

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        UPDATE caminhoes
        SET doca = %s,
            status = 'autorizado',
            autorizado_por = %s,
            horario_autorizacao = %s
        WHERE id = %s
          AND setor_doca = %s;
    """, (
        request.form.get("doca"),
        usuario,
        agora_brasil(),
        id,
        setor
    ))

    conn.commit()
    cur.close()
    conn.close()

    return redirect(url_for("encarregado"))


@app.route("/iniciar_doca/<int:id>", methods=["POST"])
def iniciar_doca(id):
    if session.get("tipo") != "encarregado":
        return redirect(url_for("login"))

    usuario = session.get("usuario")
    setor = setor_do_usuario(usuario)

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        UPDATE caminhoes
        SET status = 'na_doca',
            inicio_doca = %s
        WHERE id = %s
          AND setor_doca = %s;
    """, (
        agora_brasil(),
        id,
        setor
    ))

    conn.commit()
    cur.close()
    conn.close()

    return redirect(url_for("encarregado"))


@app.route("/finalizar_doca/<int:id>", methods=["POST"])
def finalizar_doca(id):
    if session.get("tipo") != "encarregado":
        return redirect(url_for("login"))

    usuario = session.get("usuario")
    setor = setor_do_usuario(usuario)

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        UPDATE caminhoes
        SET status = 'finalizado',
            fim_doca = %s
        WHERE id = %s
          AND setor_doca = %s;
    """, (
        agora_brasil(),
        id,
        setor
    ))

    conn.commit()
    cur.close()
    conn.close()

    return redirect(url_for("encarregado"))


@app.route("/cadastro_operacional/<int:id>", methods=["GET", "POST"])
def cadastro_operacional(id):
    if session.get("tipo") != "encarregado":
        return redirect(url_for("login"))

    usuario = session.get("usuario")
    setor = setor_do_usuario(usuario)

    conn = conectar()
    cur = conn.cursor()

    if request.method == "POST":
        cur.execute("""
            UPDATE caminhoes
            SET dts_observacao = %s,
                produtos_sku = %s,
                notas = %s,
                diferenca_os = %s,
                quantidade_nfs = %s,
                qtd_paletes_nf = %s,
                qtd_paletes_conferido = %s,
                peso_kg = %s,
                diferenca_produtos = %s,
                equipe = %s,
                operacional_cadastrado_por = %s,
                horario_operacional = %s
            WHERE id = %s
              AND setor_doca = %s;
        """, (
            request.form.get("dts_observacao"),
            request.form.get("produtos_sku"),
            request.form.get("notas"),
            request.form.get("diferenca_os"),
            request.form.get("quantidade_nfs") or None,
            request.form.get("qtd_paletes_nf") or None,
            request.form.get("qtd_paletes_conferido") or None,
            request.form.get("peso_kg") or None,
            request.form.get("diferenca_produtos"),
            request.form.get("equipe"),
            usuario,
            agora_brasil(),
            id,
            setor
        ))

        conn.commit()
        cur.close()
        conn.close()

        return redirect(url_for("encarregado"))

    cur.execute("""
        SELECT * 
        FROM caminhoes 
        WHERE id = %s
          AND setor_doca = %s;
    """, (id, setor))

    caminhao = cur.fetchone()

    cur.close()
    conn.close()

    if not caminhao:
        return redirect(url_for("encarregado"))

    return render_template(
        "cadastro_operacional.html",
        caminhao=caminhao,
        registro=caminhao,
        usuario=session.get("usuario"),
        perfil=session.get("tipo")
    )


@app.route("/admin")
def admin():
    if session.get("tipo") != "admin":
        return redirect(url_for("login"))

    busca = request.args.get("busca", "")
    data_inicio = request.args.get("data_inicio", "")
    data_fim = request.args.get("data_fim", "")

    conn = conectar()
    cur = conn.cursor()

    query = """
        SELECT *
        FROM caminhoes
        WHERE 1=1
    """

    params = []

    if busca:
        query += """
            AND (
                placa ILIKE %s OR
                motorista ILIKE %s OR
                cpf ILIKE %s OR
                empresa ILIKE %s OR
                tipo_material ILIKE %s OR
                nota_fiscal ILIKE %s OR
                setor_doca ILIKE %s OR
                doca ILIKE %s OR
                equipe ILIKE %s
            )
        """
        termo = f"%{busca}%"
        params.extend([termo, termo, termo, termo, termo, termo, termo, termo, termo])

    if data_inicio:
        query += " AND horario::date >= %s"
        params.append(data_inicio)

    if data_fim:
        query += " AND horario::date <= %s"
        params.append(data_fim)

    query += " ORDER BY id DESC;"

    cur.execute(query, params)
    caminhoes = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        "admin.html",
        caminhoes=caminhoes,
        busca=busca,
        data_inicio=data_inicio,
        data_fim=data_fim,
        usuario=session.get("usuario"),
        perfil=session.get("tipo")
    )


@app.route("/registros_encarregado")
def registros_encarregado():
    if session.get("tipo") != "admin":
        return redirect(url_for("login"))

    busca = request.args.get("busca", "")
    data_inicio = request.args.get("data_inicio", "")
    data_fim = request.args.get("data_fim", "")

    conn = conectar()
    cur = conn.cursor()

    query = """
        SELECT *
        FROM caminhoes
        WHERE horario_operacional IS NOT NULL
    """

    params = []

    if busca:
        query += """
            AND (
                placa ILIKE %s OR
                motorista ILIKE %s OR
                cpf ILIKE %s OR
                empresa ILIKE %s OR
                tipo_material ILIKE %s OR
                nota_fiscal ILIKE %s OR
                setor_doca ILIKE %s OR
                doca ILIKE %s OR
                equipe ILIKE %s
            )
        """
        termo = f"%{busca}%"
        params.extend([termo, termo, termo, termo, termo, termo, termo, termo, termo])

    if data_inicio:
        query += " AND horario_operacional::date >= %s"
        params.append(data_inicio)

    if data_fim:
        query += " AND horario_operacional::date <= %s"
        params.append(data_fim)

    query += " ORDER BY horario_operacional DESC;"

    cur.execute(query, params)
    registros = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        "registros_encarregado.html",
        registros=registros,
        busca=busca,
        data_inicio=data_inicio,
        data_fim=data_fim,
        usuario=session.get("usuario"),
        perfil=session.get("tipo")
    )


@app.route("/relatorio_excel")
def relatorio_excel():
    if session.get("tipo") != "admin":
        return redirect(url_for("login"))

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM caminhoes
        ORDER BY id DESC;
    """)

    dados = cur.fetchall()

    cur.close()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Relatorio PELOG"

    ws.append([
        "ID", "Placa", "Motorista", "CPF", "Empresa", "Material", "Nota Fiscal/DTS",
        "Setor", "Doca", "Status", "Entrada Portaria", "Autorizado Por",
        "Horário Autorização", "Início Doca", "Fim Doca",
        "DTS Observação", "Produtos SKU", "Notas", "Diferença OS",
        "Quantidade NFS", "Paletes NF", "Paletes Conferido",
        "Peso KG", "Diferença Produtos", "Equipe", "Operacional Por",
        "Horário Operacional"
    ])

    for c in dados:
        ws.append([
            c.get("id"),
            c.get("placa"),
            c.get("motorista"),
            c.get("cpf"),
            c.get("empresa"),
            c.get("tipo_material"),
            c.get("nota_fiscal"),
            c.get("setor_doca"),
            c.get("doca"),
            c.get("status"),
            c.get("horario"),
            c.get("autorizado_por"),
            c.get("horario_autorizacao"),
            c.get("inicio_doca"),
            c.get("fim_doca"),
            c.get("dts_observacao"),
            c.get("produtos_sku"),
            c.get("notas"),
            c.get("diferenca_os"),
            c.get("quantidade_nfs"),
            c.get("qtd_paletes_nf"),
            c.get("qtd_paletes_conferido"),
            c.get("peso_kg"),
            c.get("diferenca_produtos"),
            c.get("equipe"),
            c.get("operacional_cadastrado_por"),
            c.get("horario_operacional"),
        ])

    arquivo = BytesIO()
    wb.save(arquivo)
    arquivo.seek(0)

    return send_file(
        arquivo,
        as_attachment=True,
        download_name="relatorio_pelog.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/tv")
def tv():
    return render_template("tv.html")


@app.route("/dados_tv")
def dados_tv():
    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            id,
            placa,
            motorista,
            empresa,
            tipo_material,
            nota_fiscal,
            setor_doca,
            doca,
            status,
            inicio_doca
        FROM caminhoes
        WHERE status = 'na_doca'
          AND inicio_doca IS NOT NULL
        ORDER BY doca ASC, inicio_doca ASC;
    """)

    caminhoes = cur.fetchall()

    cur.close()
    conn.close()

    dados = []
    agora = agora_brasil()

    for c in caminhoes:
        tempo = 0

        if c.get("inicio_doca"):
            tempo = int((agora - c.get("inicio_doca")).total_seconds() / 60)

        dados.append({
            "id": c.get("id"),
            "placa": c.get("placa"),
            "motorista": c.get("motorista"),
            "empresa": c.get("empresa"),
            "tipo_material": c.get("tipo_material"),
            "nota_fiscal": c.get("nota_fiscal"),
            "setor_doca": c.get("setor_doca"),
            "doca": c.get("doca") or "SEM DOCA",
            "status": c.get("status"),
            "tempo": tempo
        })

    return jsonify(dados)


@app.route("/api/tv")
def api_tv():
    return dados_tv()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=True)


    