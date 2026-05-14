import pymysql
import time
import sys
import traceback
import logging
from collections import defaultdict
from datetime import datetime
import win32com.client as win32

DESTINATARIO_OPERACIONAL = ";".join([
   DESTINATARIO = "example@company.com"
])
DESTINATARIO_COMPLETO = ";".join([
    DESTINATARIO = "example@company.com"
])

ASSUNTO_BASE = "Monitoramento de Ciclos - Produção"

MODO = "operacional"
if len(sys.argv) > 1:
    MODO = sys.argv[1].lower()

DESTINATARIO = DESTINATARIO_COMPLETO if MODO == "completo" else DESTINATARIO_OPERACIONAL
ASSUNTO = f"{ASSUNTO_BASE} [{'COMPLETO' if MODO == 'completo' else 'OPERACIONAL'}]"

logging.basicConfig(
    filename="monitor_ciclos.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

def log(msg):
    print(msg)
    logging.info(msg)

def media(lista):
    return sum(lista) / len(lista) if lista else 0

conn = None
cursor = None
inicio = time.time()

try:
    log(f"INICIO | modo={MODO}")

    conn = pymysql.connect(
        host="YOUR_SQL_HOST"
        user="YOUR_SQL_USER"
        password="YOUR_SQL_PASSWORD",
        database="YOUR_DATABASE"",
        connect_timeout=15,
        cursorclass=pymysql.cursors.DictCursor
    )

    cursor = conn.cursor()

    html_email = f"""
    <html>
    <body style='font-family:Consolas,Arial,sans-serif;background:#dce6e9;padding:20px;'>

    <div style='max-width:1000px;margin:auto;background:#dce6e9;border:1px solid #c2ccd6;padding:25px;'>

    <div style='font-size:28px;font-weight:bold;color:#1f4e78;'>
    MONITORAMENTO DE CICLOS - PRODUÇÃO
    </div>

    <div style='margin-top:8px;font-size:13px;color:#666;'>
    {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
    </div>

    <hr style='margin:20px 0;border:none;border-top:2px solid #1f4e78;'>
    """

    query = """
    SELECT
        op.id AS op_id,
        op.numero AS op_numero,
        ci.created_on,
        ci.value1 AS ciclo_real,
        ci.produto AS produto_id,
        maq.nome AS maquina_nome,
        prod.codigo AS produto_codigo,
        prod.nome AS produto_nome,
        molde.codigo AS molde_codigo,
        molde.nome AS molde_nome,
        molde.mol_ciclo_medio AS ciclo_cadastrado
    FROM controle_injetora ci
    INNER JOIN ordem_producao op ON ci.ordem_producao = op.id
    LEFT JOIN maquina maq ON ci.maquina = maq.id
    LEFT JOIN produto prod ON ci.produto = prod.id
    LEFT JOIN produto molde ON ci.molde = molde.id
    WHERE
        op.status = 2
        AND ci.value1 IS NOT NULL
        AND molde.mol_ciclo_medio IS NOT NULL
    ORDER BY op.numero, ci.created_on
    """

    cursor.execute(query)
    dados = cursor.fetchall()

    ops = defaultdict(list)

    for row in dados:
        ops[row["op_numero"]].append(row)

    if ops:
        html_email += """
        <div style='font-size:22px;font-weight:bold;color:#1f4e78;margin-bottom:20px;'>
        OPS EM PRODUÇÃO
        </div>
        """

    for op_numero, registros in ops.items():

        primeiro = registros[0]
        ciclo_cadastrado = float(primeiro["ciclo_cadastrado"])
        total = len(registros)

        abaixo = 0
        acima = 0
        neutro = 0

        soma_desvio_abaixo = 0
        soma_desvio_acima = 0

        qtd_desvio_abaixo = 0
        qtd_desvio_acima = 0

        validos_media = []

        limite_superior = ciclo_cadastrado * 1.30
        limite_inferior = ciclo_cadastrado * 0.70

        for r in registros:
            ciclo = float(r["ciclo_real"])

            if ciclo < ciclo_cadastrado:
                abaixo += 1
            elif ciclo > ciclo_cadastrado:
                acima += 1
            else:
                neutro += 1

            if limite_inferior <= ciclo <= limite_superior:
                validos_media.append(ciclo)

                if ciclo < ciclo_cadastrado:
                    soma_desvio_abaixo += (ciclo_cadastrado - ciclo)
                    qtd_desvio_abaixo += 1

                elif ciclo > ciclo_cadastrado:
                    soma_desvio_acima += (ciclo - ciclo_cadastrado)
                    qtd_desvio_acima += 1

        ciclo_medio = media(validos_media)

        perc_abaixo = (abaixo / total) * 100 if total else 0
        perc_acima = (acima / total) * 100 if total else 0
        perc_neutro = (neutro / total) * 100 if total else 0

        desvio_medio_abaixo = (
            soma_desvio_abaixo / qtd_desvio_abaixo
            if qtd_desvio_abaixo else 0
        )

        desvio_medio_acima = (
            soma_desvio_acima / qtd_desvio_acima
            if qtd_desvio_acima else 0
        )

        ultimos_30 = registros[-30:]

        abaixo_30 = 0
        acima_30 = 0
        neutro_30 = 0

        for r in ultimos_30:
            ciclo = float(r["ciclo_real"])

            if ciclo < ciclo_cadastrado:
                abaixo_30 += 1
            elif ciclo > ciclo_cadastrado:
                acima_30 += 1
            else:
                neutro_30 += 1

        html_email += f"""
        <div style='
            background:#dce6e9;
            border:1px solid #c2ccd6;
            padding:18px;
            margin-bottom:18px;
        '>

        <div style='font-size:20px;font-weight:bold;color:#1f4e78;'>
        OP {op_numero}
        </div>

        <div style='margin-top:10px;line-height:1.7;font-size:14px;background:#dce6e9;'>
        <b>Produto:</b> {primeiro['produto_codigo']} / {primeiro['produto_nome']}<br>
        <b>Molde:</b> {primeiro['molde_codigo']} / {primeiro['molde_nome']}<br>
        <b>Máquina:</b> {primeiro['maquina_nome']}
        </div>

        <div style='margin-top:16px;font-size:15px;font-weight:bold;color:#1f4e78;background:#dce6e9;'>
        BLOCO 1 — PERFORMANCE GERAL
        </div>

        <div style='margin-top:10px;line-height:1.9;font-size:14px;background:#dce6e9;'>
        <b>Ciclo cadastrado:</b> {ciclo_cadastrado:.2f}s<br>
        <b>Ciclo médio real:</b> {ciclo_medio:.2f}s<br><br>

        🟢 <b>Abaixo:</b> {perc_abaixo:.1f}% ({abaixo} de {total}) | desvio médio -{desvio_medio_abaixo:.2f}s<br>
        🔴 <b>Acima:</b> {perc_acima:.1f}% ({acima} de {total}) | desvio médio +{desvio_medio_acima:.2f}s<br>
        🔵 <b>Neutro:</b> {perc_neutro:.1f}% ({neutro} de {total})
        </div>

        <div style='margin-top:16px;font-size:15px;font-weight:bold;color:#1f4e78;background:#dce6e9;'>
        BLOCO 2 — ÚLTIMOS 30 CICLOS
        </div>

        <div style='margin-top:10px;line-height:1.9;font-size:14px;background:#dce6e9;'>
        🟢 <b>Abaixo:</b> {abaixo_30}<br>
        🔴 <b>Acima:</b> {acima_30}<br>
        🔵 <b>Neutro:</b> {neutro_30}
        </div>

        </div>
        """

    if MODO == "completo":

        query_fechadas = """
        SELECT
            op.id AS op_id,
            op.numero AS op_numero,
            ci.produto AS produto_id,
            ci.value1 AS ciclo_real,
            maq.nome AS maquina_nome,
            prod.codigo AS produto_codigo,
            prod.nome AS produto_nome,
            molde.codigo AS molde_codigo,
            molde.nome AS molde_nome,
            molde.mol_ciclo_medio AS ciclo_cadastrado
        FROM controle_injetora ci
        INNER JOIN ordem_producao op
            ON ci.ordem_producao = op.id
        LEFT JOIN maquina maq
            ON ci.maquina = maq.id
        LEFT JOIN produto prod
            ON ci.produto = prod.id
        LEFT JOIN produto molde
            ON ci.molde = molde.id
        WHERE
            op.status IN (3, 4)
            AND op.termino_real >= NOW() - INTERVAL 24 HOUR
            AND ci.value1 IS NOT NULL
            AND molde.mol_ciclo_medio IS NOT NULL
        ORDER BY op.numero DESC
        """

        cursor.execute(query_fechadas)
        dados_fechadas = cursor.fetchall()

        ops_fechadas = defaultdict(list)

        for row in dados_fechadas:
            ops_fechadas[row["op_numero"]].append(row)

        if ops_fechadas:
            html_email += """
            <hr style='margin:28px 0;border:none;border-top:2px solid #1f4e78;'>

            <div style='font-size:22px;font-weight:bold;color:#1f4e78;margin-bottom:20px;background:#dce6e9;'>
            FECHAMENTO ÚLTIMAS 24 HORAS
            </div>
            """

        for op_numero, registros in ops_fechadas.items():

            primeiro = registros[0]
            ciclo_cadastrado = float(primeiro["ciclo_cadastrado"])

            limite_superior = ciclo_cadastrado * 1.30
            limite_inferior = ciclo_cadastrado * 0.70

            validos_op = []

            for r in registros:
                ciclo = float(r["ciclo_real"])

                if limite_inferior <= ciclo <= limite_superior:
                    validos_op.append(ciclo)

            ciclo_medio_op = media(validos_op)

            query_historico = """
            SELECT
                value1,
                ordem_producao
            FROM controle_injetora
            WHERE
                produto = %s
                AND ordem_producao <> %s
                AND created_on >= NOW() - INTERVAL 90 DAY
                AND value1 IS NOT NULL
            """

            cursor.execute(
                query_historico,
                (primeiro["produto_id"], primeiro["op_id"])
            )

            historico = cursor.fetchall()

            validos_hist = []
            ops_hist = set()

            for h in historico:
                ciclo = float(h["value1"])

                if limite_inferior <= ciclo <= limite_superior:
                    validos_hist.append(ciclo)
                    ops_hist.add(h["ordem_producao"])

            ciclo_medio_hist = media(validos_hist)

            html_email += f"""
            <div style='
                background:#dce6e9;
                border:1px solid #c2ccd6;
                padding:18px;
                margin-bottom:18px;
            '>

            <div style='font-size:20px;font-weight:bold;color:#1f4e78;background:#dce6e9;'>
            OP {op_numero}
            </div>

            <div style='margin-top:10px;line-height:1.7;font-size:14px;background:#dce6e9;'>
            <b>Produto:</b> {primeiro['produto_codigo']} / {primeiro['produto_nome']}<br>
            <b>Molde:</b> {primeiro['molde_codigo']} / {primeiro['molde_nome']}<br>
            <b>Máquina:</b> {primeiro['maquina_nome']}
            </div>

            <div style='margin-top:16px;font-size:15px;font-weight:bold;color:#1f4e78;background:#dce6e9;'>
            FECHAMENTO OPERACIONAL
            </div>

            <div style='margin-top:10px;line-height:1.9;font-size:14px;background:#dce6e9;'>
            <b>Ciclo cadastrado:</b> {ciclo_cadastrado:.2f}s<br>
            <b>Ciclo médio real OP:</b> {ciclo_medio_op:.2f}s<br>
            <b>Ciclo médio histórico:</b> {ciclo_medio_hist:.2f}s<br>
            <b>Quantidade de OPs analisadas:</b> {len(ops_hist)}
            </div>

            </div>
            """

    html_email += """
    </div>
    </body>
    </html>
    """

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = DESTINATARIO
    mail.Subject = ASSUNTO
    mail.HTMLBody = html_email
    mail.Send()

    fim = time.time()
    log(f"SUCESSO | tempo={fim - inicio:.2f}s")

    print("Email enviado com sucesso.")
    print(f"Tempo total: {fim - inicio:.2f}s")

except Exception:
    erro = traceback.format_exc()
    logging.error(erro)
    print(erro)

finally:
    if cursor:
        cursor.close()
    if conn:
        conn.close()