import gspread
import pyodbc

# emailAPiGoogle = python-connect-sheets@pythonsheets-344316.iam.gserviceaccount.com

# CREDENCIAIS #
print("Conexao bem Sucedida")

cursor = conn.cursor()
cursor.execute("""

                    SELECT DISTINCT

                    IRIS.Cliente				AS 'Cod. Iris',
                    PATRI.DS_PARTICAO			AS 'Nome Cliente',
                    CONTRATOS.NR_CONTRATO		AS 'Contrato Gregor',
                    AGRUP.DS_AGRUPAMENTO		AS 'CARTEIRA',
                    SERVICO.CD_SERVICO			AS 'Cód. Serviço',
                    DESC_SERVICO.DS_PRODUTO		AS 'Descrição Serviço',
                    CAST(SERVICO.VR_TOTAL AS VarChar(10)) AS 'Valor Serviço',
                    IRIS.Cidade				    AS 'CIDADE',
                    FORMAT ( CONTRATO_C.DT_INICIO_VIG, 'dd/MM/yyyy') AS 'Inicio Contrato',
                    ENTIDADES.NM_RAZAOSOC		AS 'Vendor'

                    FROM LKSATHENA.IrisSQL.dbo.Clientes AS IRIS

                    INNER JOIN PATRIMONIO_PARTICAO		AS PATRI			ON PATRI.ID_PARTICAO = IRIS.IdUnico
                    INNER JOIN CONTRATOS_CONT_PATRI		AS CONTRATOS		ON CONTRATOS.ID_PATRIMONIO = PATRI.ID_PATRIMONIO
                    INNER JOIN CONTRATOS_CONT			AS CONTRATO_C		ON CONTRATO_C.NR_CONTRATO = CONTRATOS.NR_CONTRATO
                    INNER JOIN CONTRATOS_CONT_SERVICE	AS SERVICO			ON SERVICO.NR_CONTRATO = CONTRATOS.NR_CONTRATO
                    INNER JOIN MATERIAIS				AS DESC_SERVICO		ON DESC_SERVICO.CD_PRODUTO = SERVICO.CD_SERVICO
                    INNER JOIN AGRUPAMENTO_CONTRATO		AS AGRUP			ON AGRUP.CD_AGRUPAMENTO = CONTRATO_C.CD_AGRUPAMENTO
                    LEFT JOIN ENTIDADES					AS ENTIDADES		ON ENTIDADES.CD_ENTIDADE = CONTRATO_C.CD_VENDEDOR

                    --ORDER BY IRIS.IdUnico ASC 

                """)  # Executing a query

cliente = []
for row in cursor:  # Looping over returned rows and printing them
    # my_list = [elem for elem in row]
    cliente.append([elem for elem in row])

# print(cliente)
CODE = 'API DIRECIONAMENTO'
DICT = {}

gc = gspread.service_account(filename='key.json')
sh = gc.open_by_key(CODE)
ws = sh.worksheet('GREGOR')
# row = len(ws.get_all_records())+2
# print(x)

sh.values_update(
    'GREGOR!A2',
    params={
        'valueInputOption': 'RAW'
    },
    body={
        'values': cliente
    }
)

print("Base Gregor Ouvidoria Atualizado com sucesso! ")