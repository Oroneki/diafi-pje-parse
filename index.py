import sqlite3
import os
import pathlib
import sys
import json
import pprint
import openpyxl

DB_PATH = os.path.abspath(
    str(pathlib.Path(os.getcwd()) / ".." / "diafi" / "novo.sqlite3")
)

print(DB_PATH)
if not os.path.exists(DB_PATH):
    print("ERRO: Arquivo da base de dados não encontrado.")
    sys.exit(1)
conn = sqlite3.connect(DB_PATH)
curr = conn.cursor()
curr.execute("""SELECT * FROM kv;""")

wb = openpyxl.Workbook()
ws_processos = wb.active
ws_processos.title = "Processos"
ws_movimentacoes = wb.create_sheet("Movimentações")

ws_processos.append(
    [
        "Processo",
        "Página",
        "Posição",
        "Prescrição Intercorrente",
        "Tipo",
        "Assuntos",
        "Ultimo",
        "AT CNPJ",
        "d",
        "ATIVO",
        "PASSIVO",
        "NUM",
        "AUTUACAO",
        "DISTRIBUICAO",
        "CLASSE",
        "VARA",
        "VALOR CAUSA",
        "USUARIO",
        "APLICACAO",
        "EVENTO",
        "REFERENCIA",
        "PRIORIDADE",
        "data info",
    ]
)

pi_count = 0

while True:
    subset = curr.fetchmany(50)
    if len(subset) == 0:
        break
    for (k, v) in subset:
        print(k)
        j = json.loads(v)
        ehPI = False
        infoExec = ""
        for linha in j["json"]["movimentacoes"]:
            if ("Processo Suspenso por Execução Frustrada" in linha[0]) and (
                "5 anos" in linha[1]
            ):
                ehPI = True
                pi_count = pi_count + 1
                ws_movimentacoes.append([k, "Prescrição Intercorrente", *linha])
                infoExec = linha[0].split(" - ")[0]
                continue
            mov = [k, "" if ehPI else "", *linha]
            ws_movimentacoes.append(mov)
        pi = "Suspenso - Exec. Frustrada - Arq 5 anos - " + infoExec if ehPI else ""
        infos = {}
        for [t, val] in j["json"]["infos"]:
            infos[t] = val
        processo = [
            k,
            j["infoObj"].get("pageNum", ""),
            j["infoObj"].get("index", 0),
            pi,
            j["infoObj"]["linha"][3],
            j["infoObj"]["linha"][4].replace("[", "").replace("]", ""),
            j["infoObj"]["linha"][5],
            j["infoObj"]["linha"][6],
            j["infoObj"]["linha"][7],
            j["infoObj"]["linha"][8],
            j["infoObj"]["linha"][9],
            infos["Número processo"],
            infos["Data de autuação"],
            infos["Data de distribuição"],
            infos["Classe judicial"],
            infos["Órgão julgador"],
            infos["Valor da causa"],
            infos["Usuário cadastro"],
            infos["Aplicação"].replace("\n", "-"),
            infos["Evento (s) de outro sistema"].replace("\t", "-"),
            infos.get("Processo referência", ""),
            infos.get("Prioridade", ""),
            j["infoObj"].get("datetime", ""),
        ]
        ws_processos.append(processo)

wb.save(filename="test.xlsx")
conn.close()
print(pi_count, "processos de Prescrição Intercorrente")
print("salvo arquivo do excel.")

