from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file
import os
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "pelog_secret_key")

DATABASE_URL = os.environ.get("DATABASE_URL")

if not DATABASE_URL:
    DATABASE_URL = "dbname=pelog user=postgres password=123456 host=localhost port=5432"


def conectar():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)


USUARIOS = {
    "admin": {"senha": "admin123", "tipo": "admin"},
    "portaria": {"senha": "portaria123", "tipo": "portaria"},
    "pedro_cd": {"senha": "123", "tipo": "encarregado"},
    "jean_cd": {"senha": "123", "tipo": "encarregado"},
    "jalison_cross": {"senha": "123", "tipo": "encarregado"},
    "elaine_cross": {"senha": "123", "tipo": "encarregado"},
}


def criar_colunas():
    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS caminhoes (
            id SERIAL PRIMARY KEY,
            placa VARCHAR(20),
            motorista VARCHAR(100),
            cpf VARCHAR(20),
            empresa VARCHAR(100),
            tipo_material VARCHAR(100),
            nota_fiscal VARCHAR(50),
            horario TIMESTAMP,

            doca VARCHAR(50),
            status VARCHAR(50),
            autorizado_por VARCHAR(100),
            horario_autorizacao TIMESTAMP,

            inicio_doca TIMESTAMP,
            fim_doca TIMESTAMP,

            dts_observacao TEXT,
            produtos_sku VARCHAR(255),
            notas VARCHAR(255),
            diferenca_os VARCHAR(255),
            quantidade_nfs INTEGER,
            qtd_paletes_nf INTEGER,
            qtd_paletes_conferido INTEGER,
            peso_kg NUMERIC(10,2),
            diferenca_produtos TEXT,
            operacional_cadastrado_por VARCHAR(100),
            horario_operacional TIMESTAMP
        );
    """)

    conn.commit()
    cur.close()
    conn.close()


@app.before_request
def iniciar():
    criar_colunas()


@app.route("/")
def index():
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("usuario")
        senha = request.form.get("senha")

        if usuario in USUARIOS and USUARIOS[usuario]["senha"] == senha:
            session["usuario"] = usuario
            session["tipo"] = USUARIOS[usuario]["tipo"]

            if session["tipo"] == "admin":
                return redirect(url_for("admin"))
            elif session["tipo"] == "portaria":
                return redirect(url_for("portaria"))
            elif session["tipo"] == "encarregado":
                return redirect(url_for("encarregado"))

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

    return render_template("portaria.html", caminhoes=caminhoes)


@app.route("/api/portaria")
def api_portaria():
    if session.get("tipo") != "portaria":
        return jsonify([])

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        SELECT id, placa, motorista, empresa, tipo_material, status, doca
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

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO caminhoes
        (placa, motorista, cpf, empresa, tipo_material, nota_fiscal, horario, status)
        VALUES (%s, %s, %s, %s, %s, %s, %s, 'aguardando');
    """, (
        request.form.get("placa"),
        request.form.get("motorista"),
        request.form.get("cpf"),
        request.form.get("empresa"),
        request.form.get("tipo_material"),
        request.form.get("nota_fiscal"),
        datetime.now()
    ))

    conn.commit()
    cur.close()
    conn.close()

    return redirect(url_for("portaria"))


@app.route("/encarregado")
def encarregado():
    if session.get("tipo") != "encarregado":
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

    return render_template("encarregado.html", caminhoes=caminhoes)


@app.route("/autorizar/<int:id>", methods=["POST"])
def autorizar(id):
    if session.get("tipo") != "encarregado":
        return redirect(url_for("login"))

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        UPDATE caminhoes
        SET doca = %s,
            status = 'autorizado',
            autorizado_por = %s,
            horario_autorizacao = %s
        WHERE id = %s;
    """, (
        request.form.get("doca"),
        session.get("usuario"),
        datetime.now(),
        id
    ))

    conn.commit()
    cur.close()
    conn.close()

    return redirect(url_for("encarregado"))


@app.route("/iniciar_doca/<int:id>", methods=["POST"])
def iniciar_doca(id):
    if session.get("tipo") != "encarregado":
        return redirect(url_for("login"))

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        UPDATE caminhoes
        SET status = 'na_doca',
            inicio_doca = %s
        WHERE id = %s;
    """, (
        datetime.now(),
        id
    ))

    conn.commit()
    cur.close()
    conn.close()

    return redirect(url_for("encarregado"))


@app.route("/finalizar_doca/<int:id>", methods=["POST"])
def finalizar_doca(id):
    if session.get("tipo") != "encarregado":
        return redirect(url_for("login"))

    conn = conectar()
    cur = conn.cursor()

    cur.execute("""
        UPDATE caminhoes
        SET status = 'finalizado',
            fim_doca = %s
        WHERE id = %s;
    """, (
        datetime.now(),
        id
    ))

    conn.commit()
    cur.close()
    conn.close()

    return redirect(url_for("encarregado"))


@app.route("/cadastro_operacional/<int:id>", methods=["GET", "POST"])
def cadastro_operacional(id):
    if session.get("tipo") != "encarregado":
        return redirect(url_for("login"))

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
                operacional_cadastrado_por = %s,
                horario_operacional = %s
            WHERE id = %s;
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
            session.get("usuario"),
            datetime.now(),
            id
        ))

        conn.commit()
        cur.close()
        conn.close()

        return redirect(url_for("encarregado"))

    cur.execute("SELECT * FROM caminhoes WHERE id = %s;", (id,))
    caminhao = cur.fetchone()

    cur.close()
    conn.close()

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
                doca ILIKE %s
            )
        """
        termo = f"%{busca}%"
        params.extend([termo, termo, termo, termo, termo, termo, termo])

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
                doca ILIKE %s
            )
        """
        termo = f"%{busca}%"
        params.extend([termo, termo, termo, termo, termo, termo, termo])

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
        "ID", "Placa", "Motorista", "CPF", "Empresa", "Material", "NF",
        "Doca", "Status", "Entrada Portaria", "Autorizado Por",
        "Horário Autorização", "Início Doca", "Fim Doca",
        "DTS Observação", "Produtos SKU", "Notas", "Diferença OS",
        "Quantidade NFS", "Paletes NF", "Paletes Conferido",
        "Peso KG", "Diferença Produtos", "Operacional Por",
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
            doca,
            status,
            inicio_doca,
            EXTRACT(EPOCH FROM (NOW() - inicio_doca)) / 60 AS tempo
        FROM caminhoes
        WHERE status = 'na_doca'
          AND inicio_doca IS NOT NULL
        ORDER BY doca ASC, inicio_doca ASC;
    """)

    caminhoes = cur.fetchall()

    cur.close()
    conn.close()

    dados = []

    for c in caminhoes:
        dados.append({
            "id": c.get("id"),
            "placa": c.get("placa"),
            "motorista": c.get("motorista"),
            "empresa": c.get("empresa"),
            "tipo_material": c.get("tipo_material"),
            "nota_fiscal": c.get("nota_fiscal"),
            "doca": c.get("doca") or "SEM DOCA",
            "status": c.get("status"),
            "tempo": int(c.get("tempo")) if c.get("tempo") else 0
        })

    return jsonify(dados)


@app.route("/api/tv")
def api_tv():
    return dados_tv()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=True)