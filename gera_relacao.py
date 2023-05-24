# Relacao de escritorio
import pandas as pd

# Função para ler os dados
def ler_dados():
    opcao = '_'
    while (opcao not in ('1', '2')):
        opcao = input("1 - Já tenho o banco completo \n2 - Quero que o programa una os csv's individuais\n\n")
    if opcao == '1':
        try:
            banco = input('Qual o nome do seu banco (EX: Exportado_2_3_123.csv)\n\n')
            dados = pd.read_csv(banco)
        except:
            try:
                import pandas as pd
                dados = pd.read_csv(banco)
            except:
                from glob import glob
                arquivos = glob('*.csv')
                print(f'\n\nOs arquivos na sua pasta são {arquivos}, reveja o nome do arquivo digitado')
                while (banco not in arquivos):
                    banco = input('Qual o nome do seu banco (EX: Exportado_2_3_123.csv)\n\n')
                dados = pd.read_csv(banco)
        else:
            input(
                "Coloque todos csv's na pasta exportados, deixe apenas os arquivos fornecidos pelo SIPEIA, aperte <ENTER> para continuar")
            from glob import glob
            arquivos = glob('./exportados/*.csv')
            for i in range(len(arquivos)):
                path = arquivos[i]
                if (i == 0):
                    dados = pd.read_csv(path)
                else:
                    df = pd.read_csv(path)
                    dados = pd.concat([dados, df], axis=0)

    dados = dados[['CNPJ', 'Razão Social', 'Pesquisa', 'Modelo', 'Status da Empresa', 'Escritório do Contador', 'FAC']]

    dados.fillna(' ', inplace=True)
    dados.columns = ['CNPJ', 'Razão Social', 'Questionário', 'Modelo', 'Já entregue', 'Escritório', 'FAC']

    # Correção das FAC's
    facs = dados[dados['FAC'].isin([8, 12, 13])].copy()
    facs.replace('Fac*', 'Não', inplace=True)
    for i in facs.index:
        dados.loc[i] = facs.loc[i]

    dados['Já entregue'] = dados['Já entregue'].replace(
        ['NADA FEITO', 'ABORDAGEM EM ANDAMENTO', 'ABORDADA', 'ACORDADA', 'EM ATRASO', 'RENEGOCIADA', 'COBRADA',
         'NOTIFICADA', 'COLETADA', 'FAC'], ['Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Não', 'Sim', 'FAC*'])
    return (dados)


# Main
opcao = '_'

print('Você quer a relação para:')
while (opcao not in ['1', '2']):
    opcao = input('1 - Um escritório\n2 - Todos escritórios\n\n')

if (opcao == '1'):

    dados = ler_dados()
    escritorio = input('Qual o escritório buscado? (Digite o nome igual o do SIPEIA)\n\n')

    if ('/' in escritorio):
        esc_novo = escritorio.replace('/', '-')
        dados['Escritório'].replace(escritorio, esc_novo, inplace=True)
        escritorio = esc_novo

    while (escritorio not in dados['Escritório'].unique()):
        print('Escritório não encontrado, verifique no SIPEIA e digite novamente')
        escritorio = input('Qual o escritório buscado?\n\n')

    dados.query('Escritório == @escritorio', inplace=True)

    if ('FAC Concluída*' in dados['Já entregue']):
        dados = pd.concat((dados,
                           pd.DataFrame(['* FAC Concluída: Empresa foi retirada da pesquisa neste ano'], columns=['CNPJ'])),
            ignore_index=True, axis=0)

    dados.fillna(' ', inplace=True)
    dados.to_excel('Lista Econômicas Anuais.xlsx', index=False)

elif (opcao == '2'):
    dados = ler_dados()

    for esc in dados['Escritório'].unique():

        if ('/' in esc):
            esc_novo = esc.replace('/', '-')
            dados['Escritório'].replace(esc, esc_novo, inplace=True)
            esc = esc_novo

        dados_temp = dados.query("Escritório == @esc")

        if ('FAC Concluída*' in dados['Já entregue']):
            dados_temp = pd.concat((dados, pd.DataFrame(['* FAC Concluída: Empresa foi retirada da pesquisa neste ano'],
                                                        columns=['CNPJ'])), ignore_index=True, axis=0)

        dados_temp.fillna(' ', inplace=True)
        dados_temp.to_excel(esc + ' - Econômicas Anuais.xlsx', index=False)