from pandas import read_excel
nome_promocao = input("Informe o nome da promoção/arquivo: ")
try:
    planilha = read_excel(f"{nome_promocao}.xlsx", skiprows=3)
    hierarquia = planilha.columns[1]

    if hierarquia == 'Estilo':
        coluna = 'cod_niv0'
    elif hierarquia == 'Subsetor':
        coluna = 'cod_niv2'
    elif hierarquia == 'Seção':
        coluna = 'cod_niv4'

    planilha.insert(column='des_selecao', loc=1, value=f'{nome_promocao}')
    planilha.rename(columns={hierarquia: coluna}, inplace=True)
    v = planilha.to_csv(f'{nome_promocao}.csv', index=False,
                        sep=';', columns=[coluna, 'des_selecao'])
    planilha2 = read_excel(f"{nome_promocao}.xlsx", sheet_name='Lojas')
    planilha2.rename(columns={'Loja': 'cod_unidade'}, inplace=True)
    v2 = planilha2.to_csv(f'{nome_promocao}_lojas.csv',
                          index=False, sep=';', columns=['cod_unidade'])
except FileNotFoundError:
    print("Arquivo informado não encontrado!")
except ValueError:
    print("Arquivo de lojas não encontrado, gerado apenas itens!")
