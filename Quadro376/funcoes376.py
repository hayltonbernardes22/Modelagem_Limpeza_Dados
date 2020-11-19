class automacao376:

    def import_txt():
        import pandas as pd
        from tkinter import Tk
        from tkinter import filedialog

        # Importando Arquivo de layout
        path_layout = 'C:\\Users\\haylton.neto\\OneDrive - GRANT THORNTON BRASIL\\Projetos_Digital_em_andamento\\Proj_Python_Versao1\\parametros\\LQ376.xlsx'
        layout = pd.read_excel(path_layout)      ### leitura do arquivo com o layout atualizado do quadro 376

        class qe_s:
        
            def __init__(self):
                self.df = {nome[s]: [] for s in range(nome.shape[0])}
    
            def make_df(self):
                lines = []
                root = Tk()
                root.attributes("-topmost", True)
                root.withdraw()
                for file in filedialog.askopenfilenames(filetypes=[("Arquivos de Text", "*.txt")]):
                    for line in open(file, encoding="utf-8"):
                        line = line.replace(",",".").strip()
                        for x in range(nome.shape[0]):
                            self.df[nome[x]].append(line[i[x]:j[x]])
                        
                return pd.DataFrame(self.df)

        nome     = layout.Campo                      ### guarda os nomes das colunas existentes no quadro 376
        i        = layout.Ind_inf                    ### guarda a posição no arquivo que inicia um determinado campo
        j        = layout.Ind_sup                    ### guarda a posição no arquivo que termina um determinado campo
        quadro = qe_s()       
        df_main = quadro.make_df()

        return df_main

        ###########################################
        ######## Criticas 376 (Validacoes) ########
        ###########################################

    def valida_criticas(df_main):
        import pandas as pd
        from datetime import date, datetime, timedelta
        
        ###################################
        ###Importacao do Quadro376.xlsx ###
        ####que serão usados para fins#####
        ######### de comparação ###########
        ###################################

        path_parametro = 'C:\\Users\\haylton.neto\\OneDrive - GRANT THORNTON BRASIL\\Projetos_Digital_em_andamento\\Proj_Python_Versao1\\parametros\\Quadro376.xlsx'
        # Pasta Cod_SUSEP
        arq_entcod = pd.read_excel(path_parametro,
                                    'Cod_SUSEP',fileEncoding = "UTF-8",encoding = 'latin_1', dtype=str)

        # Pasta TPMOID
        arq_tpmoid = pd.read_excel(path_parametro,
                                    'TPMOID',fileEncoding = "UTF-8",encoding = 'latin_1', dtype=str)

        # Pasta rel_CMPID_TPMOID
        arq_relacao = pd.read_excel(path_parametro,
                                    'rel_CMPID_TPMOID',fileEncoding = "UTF-8",encoding = 'latin_1', dtype=str)

        # Pasta CMPID
        arq_cmpid = pd.read_excel(path_parametro,
                                    'CMPID',fileEncoding = "UTF-8",encoding = 'latin_1', dtype=str)

        # Pasta RAMCODIGO
        arq_ramcod = pd.read_excel(path_parametro,
                                    'RAMCODIGO',fileEncoding = "UTF-8",encoding = 'latin_1', dtype=str)

        # Arquivo Modelo Empresas
        arq_codsusep = pd.read_excel(path_parametro,
                                        'Cod_SUSEP',fileEncoding = "UTF-8",encoding = 'latin_1', dtype=str) # Lendo arquivo parametro

        # Criando uma lista para guardar todas as criticas impeditivas
        criticas_impeditivas = []
        criticas = []

        #7392.2 Verifica o tamanho padrão da linha (deve conter 126 caracteres)
        tamanho = 0
        for i in range(0,df_main.shape[0]):                    #Passando por cada linha
            for iten in df_main.columns:                       #Passando por cada coluna
                tamanho = tamanho + len(df_main[iten][i])      #Contando Caracteres por linha
            if tamanho != 126:
                critica7392_2 = ("7392.2", "Erro na linha {} o número de caracteres é de: {}".format(i,tamanho))
                criticas_impeditivas.append(critica7392_2)
            tamanho = 0

        #7392.4 Verifica se o campo ENTCODIGO corresponde à sociedade que está enviando o FIP/SUSEP
        entcod = set(arq_entcod['Cod_SUSEP'].astype('str'))                   #Escolhendo a coluna do arquivo modelo
        dadosentcod = set(df_main['ENTCODIGO'])                               #Escolhendo a coluna dos dados
        for itemtentcod in dadosentcod: 
            if itemtentcod not in entcod:
                critica7392_4 = ('7392.4','ENTCODIGO {} nao corresponde a uma operacao valida'.format(itemtentcod))
                criticas_impeditivas.append(critica7392_4)

        #7392.5 Verifica se o campo MRFMESANO corresponde, respectivamente, ao ano, mês e último dia do mês de referência do FIP/SUSEP
        df_main.loc[:,'DTMRFMESANO'] = df_main['MRFMESANO'].astype('datetime64')     #Criando uma nova coluna no formato Date
        # Criando funcao last_day_of_month
        import datetime
        def last_day_of_month(any_day):
            next_month = any_day.replace(day=28) + datetime.timedelta(days=4)  # this will never fail
            return next_month - datetime.timedelta(days=next_month.day)
        # Selecionando dadas validas
        dadosmrfmesano = set(df_main['DTMRFMESANO'])
        # Criando e adicionando anos na lista
        years = []
        for data in dadosmrfmesano:
            years.append(int(data.strftime("%Y")))
        datas_validas = []
        for year in years:
            for month in range(1,13):
                x = (last_day_of_month(datetime.date(year, month, 1)))
                datas_validas.append(x)
        datas_validas = pd.DataFrame(datas_validas)
        datas_validas = datas_validas.astype('datetime64')

        # Realizando o check
        mrfmesano = set(datas_validas[0].astype('datetime64'))                 #Escolhendo a coluna do arquivo modelo
        dadosmrfmesano = set(df_main['DTMRFMESANO'].astype('datetime64'))                                 #Escolhendo a coluna dos dados
        for itemmrfmesano in dadosmrfmesano: 
            if itemmrfmesano not in mrfmesano:
                critica7392_5 = ('7392.5', 'MRFMESANO {} nao corresponde ao ano, mês e último dia do mês de referência'.format(itemmrfmesano))
                criticas_impeditivas.append(critica7392_5)

        #7392.6 Verifica se o campo QUAID corresponde ao quadro 376
        array6 = set(df_main['QUAID'])           ##Verificando todos os Quadros que estamos tratando no documento.
        for itemarray6 in array6:
            if itemarray6 != '376':
                critica7392_6 = ('7392.6', 'Este esta se tratando do(s) Quadro(s):{}'.format(itemarray6))
                criticas_impeditivas.append(critica7392_6)

        #7392.7 Verifica se o campo TPMOID corresponde a um tipo de movimento válido
        tpmoid = set(arq_tpmoid['TPMOID'].astype('int64'))                                    #Escolhendo a coluna do arquivo modelo
        dadostpmoid = set(df_main['TPMOID'].astype('int64'))                                                #Escolhendo a coluna dos dados
        for itemtpmoid in dadostpmoid: 
            if itemtpmoid not in tpmoid:
                critica7392_7 = ('7392.7', 'TPMOID {} nao corresponde a uma operacao valida'.format(itemtpmoid))
                criticas_impeditivas.append(critica7392_7)

        #7392.8 Valida a correspondência entre os campos TPMOID e CMPID
        mod_relacao = set(arq_relacao['relacao'].astype('str'))
        relacao = [df_main['CMPID'].astype('str') + df_main['TPMOID'].astype('str')]
        for itemrelacao in relacao:
            relacao2 = set(itemrelacao)
            for itemrelacao2 in relacao2:
                if itemrelacao2 not in mod_relacao:
                    itemrelacao2 = itemrelacao2
                    critica7392_8 = ('7392.8', 'A relacao {} (CMPID {} /TPMOID {} ) nao é uma relacao possivel.'.format(itemrelacao2,itemrelacao2[:4],itemrelacao2[4:]))
                    criticas_impeditivas.append(critica7392_8)

        #7392.9 Verifica se o campo CMPID corresponde a um tipo de operação válida 
        cmpid = set(arq_cmpid['CMPID'].astype('str'))                        #Escolhendo a coluna do arquivo modelo
        dadoscmpid = set(df_main['CMPID'])                                   #Escolhendo a coluna dos dados
        for itemcmpid in dadoscmpid: 
            if itemcmpid not in cmpid:
                critica7392_9 = ('7392.9', 'CMPID {} nao corresponde a uma operacao valida'.format(itemcmpid))
                criticas_impeditivas.append(critica7392_9)

        #7392.10 Verifica se o campo RAMCODIGO corresponde, respectivamente, a um grupo de ramos e ramo válidos e operados pela companhia no mês de referência
        ramcod = set(arq_ramcod['ramo_s1'].astype('int64'))                       #Escolhendo a coluna do arquivo
        dadosramcod = set(df_main['RAMCODIGO'].astype('int64'))                                 #Escolhendo a coluna dos dados
        for itemramcod in dadosramcod: 
            if itemramcod not in ramcod:
                critica7392_10 = ('7392.10', 'RAMCODIGO {} nao corresponde a uma operacao valida'.format(itemramcod))
                criticas_impeditivas.append(critica7392_10)

        #7392.11 Verifica se o campo RAMCODIGO não foi preenchido com os ramos 0588, 0589, 0983, 0986, 0991, 0992, 0994, 0996, 1066, 1383, 1386, 1391, 1392, 1396, 1603 e 2201 
        ramcodver = set(arq_ramcod['excecoes'].astype('str'))       #Escolhendo a coluna do arquivo modelo
        dadosramcod = set(df_main['RAMCODIGO'])                                 #Escolhendo a coluna dos dados
        for itemramcod in dadosramcod: 
            if itemramcod in ramcodver:
                critica7392_11 = ('7392.11', 'RAMCODIGO {} nao corresponde a uma operacao valida'.format(itemramcod))
                criticas_impeditivas.append(critica7392_11)
            
        #7392.13 Verifica se os campos ESRDATAINICIO, ESRDATAFIM, ESRDATAOCORR, ESRDATAREG e ESRDATACOMUNICA correspondem a uma data válida 

        #####################################################
        ##Readicionando as colunas de Data de str para date##
        #####################################################

        df_main.loc[:,'MES_ano'] = df_main['DTMRFMESANO'].dt.strftime("%Y-%m").astype('datetime64') #Criando coluna MÊSANO
        df_main.loc[:,'MRFMES_ano'] = df_main['DTMRFMESANO'].dt.strftime("%Y-%m-01").astype('datetime64') #Criando coluna MÊSANO-01
        df_main.loc[:,'DTESRDATAINICIO'] = df_main['ESRDATAINICIO'].astype('datetime64')
        df_main.loc[:,'DTESRDATAFIM'] = df_main['ESRDATAFIM'].astype('datetime64')
        df_main.loc[:,'DTESRDATAOCORR'] = df_main['ESRDATAOCORR'].astype('datetime64')
        df_main.loc[:,'DTESRDATAREG'] = df_main['ESRDATAREG'].astype('datetime64')
        df_main.loc[:,'DTESRDATACOMUNICA'] = df_main['ESRDATACOMUNICA'].astype('datetime64')

        #Adicionado colunas DTESRDATAINICIO, DTESRDATAFIM, DTESRDATAOCORR, DTESRDATAREG e DTESRDATACOMUNICA no tipo Date

        #7392.12 Verifica se o valor dos campos ESRVALORMOV e ESRVALORMON é float

        ###########################################
        ##Readicionando as colunas str para float##
        ###########################################

        df_main.loc[:,'fESRVALORMOV'] = df_main['ESRVALORMOV'].astype('float64')
        df_main.loc[:,'fESRVALORMON'] = df_main['ESRVALORMON'].astype('float64')

        #Adicionado colunas fESRVALORMOV e fESRVALORMON no tipo Float

        ## Transformando criticas impeditivas para DataFrame
        from datetime import datetime
        if criticas_impeditivas == []:
            criticas_impeditivas.append('Nenhuma Critica Impeditiva encontrada no Quadro 376')
        df_criticas_i = pd.DataFrame(criticas_impeditivas)     #criacao do DataFrame das criticas impeditivas
        df_criticas_i = df_criticas_i.rename(columns={0:'ID da Critica',1:'Descricao', 2:'Codigo Invalido'})
        data_presente = datetime.now()
        nome = input("Insira seu nome: ")
        df_criticas_i.loc[:,'ID da Empresa'] = itemtentcod  ### insere a coluna ID da Empresa com a da tabela
        df_criticas_i.loc[:,'Criador'] = nome ### insere a coluna com o nome do criador da tabela
        df_criticas_i.loc[:,'Data do output'] = data_presente ### insere a coluna com a data de criacao da tabela

        ###########################################
        ## Criticas (Validacoes) Não Impeditivas ##
        ###########################################

        #7392.14 Verifica se o campo ESRCODCESS corresponde a um código de sociedade válido
        dadosesrcodcess = set(df_main['ESRCODCESS'])       #Escolhendo a coluna dos dados
        for itemtesrcodcess in dadosesrcodcess: 
            if itemtesrcodcess not in entcod:
                critica7392_14 = ('7392.14', 'ESRCODCESS {} nao corresponde a um codigo de sociedade valido'.format(itemtesrcodcess))
                criticas.append(critica7392_14)

        # 7392.15
        # Criando DataFrame copia com Campos CMPID, ESRCODCESS, ENTCODIGO no formato int
        df_mains = df_main.copy()
        df_mains['CMPID'] = df_mains['CMPID'].astype('int64')
        df_mains['ESRCODCESS'] = df_mains['ESRCODCESS'].astype('int64')
        df_mains['ENTCODIGO'] = df_mains['ENTCODIGO'].astype('int64')
        #7392.15 Valida a correspondência entre os campos CMPID e ESRCODCESS
        relacao15_1 = df_mains[(df_mains['CMPID'] == 1001) & (df_mains['ESRCODCESS'].values != df_mains['ENTCODIGO'].values)]
        relacao15_4 = df_mains[(df_mains['CMPID'] == 1004) & (df_mains['ESRCODCESS'].values != df_mains['ENTCODIGO'].values)]
        relacao15_6 = df_mains[(df_mains['CMPID'] == 1006) & (df_mains['ESRCODCESS'].values != df_mains['ENTCODIGO'].values)]
        relacao15_9 = df_mains[(df_mains['CMPID'] == 1009) & (df_mains['ESRCODCESS'].values != df_mains['ENTCODIGO'].values)]
        i = 24   #Definindo numero de colunas

        # Adicionando criticas com CMPIDs 1001, 1004, 1006, 1009 na lista 'criticas'
        if relacao15_1.shape != (0,i): 
            critica7392_15_1 = ('7392.15', 'O CMPID é 1001 e o ESRCODCESS é diferente do ENTCODIGO') 
            criticas.append(critica7392_15_1)
        elif relacao15_4.shape != (0,i): 
            critica7392_15_4 = ('7392.15', 'O CMPID é 1004 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_4)
        elif relacao15_6.shape != (0,i): 
            critica7392_15_6 = ('7392.15', 'O CMPID é 1006 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_6)
        elif relacao15_9.shape != (0,i): 
            critica7392_15_9 = ('7392.15', 'O CMPID é 1009 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_9)

        #7392.15 Valida a correspondência entre os campos CMPID e ESRCODCESS
        relacao15_2 = df_mains[(df_mains['CMPID'] == 1002) & (df_mains['ESRCODCESS'].values <= 1) & (df_mains['ESRCODCESS'].values >= 9999)] 
        relacao15_3 = df_mains[(df_mains['CMPID'] == 1003) & (df_mains['ESRCODCESS'].values <= 1) & (df_mains['ESRCODCESS'].values >= 9999)] 
        relacao15_7 = df_mains[(df_mains['CMPID'] == 1007) & (df_mains['ESRCODCESS'].values <= 1) & (df_mains['ESRCODCESS'].values >= 9999)] 
        relacao15_8 = df_mains[(df_mains['CMPID'] == 1008) & (df_mains['ESRCODCESS'].values <= 1) & (df_mains['ESRCODCESS'].values >= 9999)] 
        # Adicionando criticas com CMPIDs 1002, 1003, 1007, 1008 na lista 'criticas'
        if relacao15_2.shape != (0,i): 
            critica7392_15_2 = ('7392.15', 'O CMPID é 1002 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_2)
        elif relacao15_3.shape != (0,i): 
            critica7392_15_3 = ('7392.15', 'O CMPID é 1003 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_3)
        elif relacao15_7.shape != (0,i): 
            critica7392_15_7 = ('7392.15', 'O CMPID é 1007 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_7)
        elif relacao15_8.shape != (0,i): 
            critica7392_15_8 = ('7392.15', 'O CMPID é 1008 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_8)

        #7392.15 Valida a correspondência entre os campos CMPID e ESRCODCESS
        relacao15_12 = df_mains[(df_mains['CMPID'] == 1012) & (df_mains['ESRCODCESS'].values <= 30000) & (df_mains['ESRCODCESS'].values >= 59999)] 
        relacao15_11 = df_mains[(df_mains['CMPID'] == 1011) & (df_mains['ESRCODCESS'].values <= 30000) & (df_mains['ESRCODCESS'].values >= 59999)]
        relacao15_13 = df_mains[(df_mains['CMPID'] == 1013) & (df_mains['ESRCODCESS'].values <= 30000) & (df_mains['ESRCODCESS'].values >= 59999)]
        relacao15_14 = df_mains[(df_mains['CMPID'] == 1014) & (df_mains['ESRCODCESS'].values <= 30000) & (df_mains['ESRCODCESS'].values >= 59999)]
        relacao15_5 = df_mains[(df_mains['CMPID'] == 1005) & (df_mains['ESRCODCESS'].values <= 30000) & (df_mains['ESRCODCESS'].values >= 59999)]
        relacao15_10 = df_mains[(df_mains['CMPID'] == 1010) & (df_mains['ESRCODCESS'].values <= 30000) & (df_mains['ESRCODCESS'].values >= 59999)]
        # Adicionando criticas com CMPIDs 1005, 1010, 1011, 1012, 1013, 1014 na lista 'criticas'
        if relacao15_12.shape != (0,i):
            critica7392_15_12 = ('7392.15', 'O CMPID é 1012 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_12)
        elif relacao15_11.shape != (0,i): 
            critica7392_15_11 = ('7392.15', 'O CMPID é 1011 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_11)
        elif relacao15_13.shape != (0,i): 
            critica7392_15_13 = ('7392.15', 'O CMPID é 1013 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_13)
        elif relacao15_14.shape != (0,i): 
            critica7392_15_14 = ('7392.15', 'O CMPID é 1014 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_14)
        elif relacao15_10.shape != (0,i): 
            critica7392_15_10 = ('7392.15', 'O CMPID é 1010 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_10)
        elif relacao15_5.shape != (0,i): 
            critica7392_15_5 = ('7392.15', 'O CMPID é 1005 e o ESRCODCESS é diferente do ENTCODIGO')
            criticas.append(critica7392_15_5)

        #7392.16 Se o tipo de operação for 'direto - administrativo' ou 'cosseguro aceito - administrativo' ou 'cosseguro cedido - administrativo' ou 'recuperação de sinistros não pagos - administrativo' ou 'recuperação de sinistros já pagos - administrativo' ou 'salvados e ressarcidos - administrativo' ou ‘salvados e ressarcidos ao ressegurador - administrativo’, Recuperação de Sinistros não Pagos – Administrativo, Recuperação de Sinistros já Pagos – Administrativo verifica se o ano da data dos campos ESRDATAINICIO, ESRDATAFIM, ESRDATAOCORR, ESRDATAREG e ESRDATACOMUNICA está entre os limites de trinta  anos para mais ou para menos do ano da data do campo MRFMESANO 
        #Criando um DataFrame com colunas de Data para cpmparação. 
        #Adicionando MRFMESANO minimo e maximo
        df_dates_comparacao = pd.DataFrame(df_main['DTMRFMESANO'] - timedelta(days = 10950))
        df_dates_comparacao.loc[:,'DTMRFMESANO_max'] = df_main['DTMRFMESANO'] + timedelta(days = 10950)

        df16_i = df_main.loc[(df_main['DTESRDATAINICIO'] < df_dates_comparacao['DTMRFMESANO']) | (df_main['DTESRDATAINICIO'] > df_dates_comparacao['DTMRFMESANO_max'])] 
        if len(df16_i) > 0:
            df16_i.loc[:,'ID da Critica'] = ('7392.16')

        df16_f = df_main.loc[(df_main['DTESRDATAFIM'] < df_dates_comparacao['DTMRFMESANO']) | (df_main['DTESRDATAFIM'] > df_dates_comparacao['DTMRFMESANO_max'])] 
        if len(df16_f) > 0:
            df16_f.loc[:,'ID da Critica'] = ('7392.16')

        df16_o = df_main.loc[(df_main['DTESRDATAOCORR'] < df_dates_comparacao['DTMRFMESANO']) | (df_main['DTESRDATAOCORR'] > df_dates_comparacao['DTMRFMESANO_max'])]
        if len(df16_o) > 0:
            df16_o.loc[:,'ID da Critica'] = ('7392.16')

        df16_r = df_main.loc[(df_main['DTESRDATAREG'] < df_dates_comparacao['DTMRFMESANO']) | (df_main['DTESRDATAREG'] > df_dates_comparacao['DTMRFMESANO_max'])] 
        if len(df16_r) > 0:
            df16_r.loc[:,'ID da Critica'] = ('7392.16')

        df16_c = df_main.loc[(df_main['DTESRDATACOMUNICA'] < df_dates_comparacao['DTMRFMESANO']) | (df_main['DTESRDATACOMUNICA'] > df_dates_comparacao['DTMRFMESANO_max'])] 
        if len(df16_c) > 0:
            df16_c.loc[:,'ID da Critica'] = ('7392.16')

        #7392.17 Verifica se a data do campo ESRDATAFIM é posterior à data do campo ESRDATAINICIO 
        #Criando uma mascara para realizar a validação
        df17 = df_main.loc[(df_main['DTESRDATAFIM'] < df_main['DTESRDATAINICIO'])] #Transformando a mascara em DataFrame
        if len(df17) > 0:
            df17.loc[:,'ID da Critica'] = ('7392.17')

        #7392.18 Verifica se a data do campo ESRDATAOCORR está entre as datas dos campos ESRDATAINICIO e ESRDATAFIM 
        df18 = df_main.loc[(df_main['DTESRDATAOCORR'] < df_main['DTESRDATAINICIO']) | (df_main['DTESRDATAOCORR'] > df_main['DTESRDATAFIM'])]
        if len(df18) > 0:
            df18.loc[:,'ID da Critica'] = ('7392.18')

        #7392.19 Verifica se a data do campo ESRDATAOCORR é igual ou anterior à data dos campos ESRDATAREG e ESRDATACOMUNICA 
        df19 = df_main.loc[(df_main['DTESRDATAOCORR'] > df_main['DTESRDATAREG']) | (df_main['DTESRDATAOCORR'] > df_main['DTESRDATACOMUNICA'])]
        if len(df19) > 0:
            df19.loc[:,'ID da Critica'] = ('7392.19')

        #7392.20 Verifica se a data do campo ESRDATACOMUNICA é igual ou anterior à data do campo ESRDATAREG
        df20 = df_main.loc[(df_main['DTESRDATACOMUNICA'] > df_main['DTESRDATAREG'])]
        if len(df20) > 0:
            df20.loc[:,'ID da Critica'] = ('7392.20')

        #7392.21 Verifica se a data dos campos ESRDATAINICIO, ESRDATAOCORR, ESRDATAREG e ESRDATACOMUNICA é igual ou anterior à data do campo MRFMESANO 
        df21 = df_main.loc[(df_main['DTMRFMESANO'] < df_main['DTESRDATAINICIO']) | (df_main['DTMRFMESANO'] < df_main['DTESRDATAOCORR']) | (df_main['DTMRFMESANO'] < df_main['DTESRDATAREG']) | (df_main['DTMRFMESANO'] < df_main['DTESRDATACOMUNICA'])]
        if len(df21) > 0:
            df21.loc[:,'ID da Critica'] = ('7392.21')

        #7392.22 Se o tipo de movimento for 'aviso' e o tipo de operação for 'direto - administrativo' ou 'cosseguro aceito -administrativo' ou 'cosseguro cedido - administrativo' ou 'direto - judicial' ou 'cosseguro aceito -judicial' ou 'cosseguro cedido - judicial', verifica se o mês e ano da data do campo ESRDATAREG é igual ao mês e ano dadata do campo MRFMESANO
        mask22_tpmoid = (df_main['TPMOID'] == '0001') & ((df_main['CMPID'] == '1001')|(df_main['CMPID'] == '1002')|(df_main['CMPID'] == '1003')|(df_main['CMPID'] == '1006')|(df_main['CMPID'] == '1007')|(df_main['CMPID'] == '1008'))
        df22 = df_main.loc[(mask22_tpmoid) & (df_main['DTMRFMESANO'].dt.strftime("%Y-%m") != df_main['DTESRDATAREG'].dt.strftime("%Y-%m"))]
        if len(df22) > 0:
            df22.loc[:,'ID da Critica'] = ('7392.22')

        #7392.23 Verifica se o valor dos campos ESRVALORMOV e ESRVALORMON é igual ou maior do que zero, exceto para o tipo de movimento’ reavaliação’(0002), quando qualquer um dos campos pode assumir valor negativo.  O campo ESRVALORMON também poderá assumir valor negativo para o tipo de movimento ‘cancelamento’(0005), nos casos de estorno, e para os tipos de movimento ‘reabertura’ (0006) e ‘recuperação de sinistros – transferência de ativo redutor de PSL para crédito com ressegurador’ (0014)
        #ESRVALORMOV e ESRVALORMON não pode ser negativo para Ativação (TPMOID 0001)
        df23_1 = df_main.loc[((df_main['TPMOID'] == '0001') & (df_main['fESRVALORMOV'] < 0.0) | (df_main['fESRVALORMON'] < 0.0))]

        if len(df23_1) > 0:
            df23_1.loc[:,'ID da Critica'] = ('7392.23')

        #7392.23 ESRVALORMOV e ESRVALORMON não pode ser negativo para Recebimento Parcial  (TPMOID 0003)
        df23_3 = df_main.loc[((df_main['TPMOID'] == '0003') & (df_main['fESRVALORMOV'] < 0.0) | (df_main['fESRVALORMON'] < 0.0))]
        if len(df23_3) > 0:
            df23_3.loc[:,'ID da Critica'] = ('7392.23')

        #7392.23 ESRVALORMOV e ESRVALORMON não pode ser negativo para Recebimento Final (ou Total)  (TPMOID 0004)
        df23_4 = df_main.loc[((df_main['TPMOID'] == '0004') & (df_main['fESRVALORMOV'] < 0.0) | (df_main['fESRVALORMON'] < 0.0))]
        if len(df23_4) > 0:
            df23_4.loc[:,'ID da Critica'] = ('7392.23')

        #7392.25 Verifica se pelo menos um dos campos possui valor diferente de zero: ESRVALORMOV e ESRVALORMON
        df25 = df_main.loc[(df_main['fESRVALORMON'] == 0.0) & (df_main['fESRVALORMOV'] == 0.0)]

        if len(df25) > 0:
            df25.loc[:,'ID da Critica'] = ('7392.25')

        #7392.26 Se o tipo de movimento for 'aviso' e o tipo de operação for 'direto - administrativo' ou 'cosseguro cedido - administrativo' ou 'direto - judicial' ou 'cosseguro cedido - judicial', verifica se a diferença de dias entre a data dos campos ESRDATACOMUNICA e ESRDATAREG é de até quinze dias 
        #Adicionando ESRDATACOMUNICA + 15days
        df_dates_comparacao.loc[:,'DTESRDATACOMUNICA'] = df_main['DTESRDATACOMUNICA'] + timedelta(days = 15)
        df26 = df_main.loc[(df_main['DTESRDATAREG'] > df_dates_comparacao['DTESRDATACOMUNICA']) & (df_main['TPMOID'] == '0001')]
        if len(df26) > 0:
            df26.loc[:,'ID da Critica'] = ('7392.26')

        #7392.27 Se o valor do campo ESRVALORMOV for igual a zero e o valor do campo ESRVALORMON for diferente de zero, verifica se o tipo de movimento é ‘reavaliação’ ou 'cancelamento' (TPMOID 0002 ou 0005)
        df27 = df_main.loc[(df_main['fESRVALORMOV'] == 0.0) & (df_main['fESRVALORMON'] != 0.0) & ((df_main['TPMOID'] != '0002')&(df_main['TPMOID'] != '0005'))]
        if len(df27) > 0:
            df27.loc[:,'ID da Critica'] = ('7392.27')

        #7392.29 Se o tipo de movimento for 'liquidação parcial' ou 'liquidação final', (TPMOID 0003 e 0004)verifica se o valor do campo ESRVALORMON é igual a zero
        # Critica 7392.29 liquidação parcial
        df29_p = df_main.loc[(df_main['TPMOID'] == '0003') & (df_main['fESRVALORMON'] != 0)]
        if len(df29_p) > 0:
            df29_p.loc[:,'ID da Critica'] = ('7392.29')

        # Critica 7392.29 liquidação final
        df29_f = df_main.loc[(df_main['TPMOID'] == '0004') & (df_main['fESRVALORMON'] != 0)]
        if len(df29_f) > 0:
            df29_f.loc[:,'ID da Critica'] = ('7392.29')

        ###################################
        # Inicio das tratativas do output #
        ###################################
        ent = []                                          # Criando lista vazia
        for i in arq_codsusep['Cod_SUSEP'].astype('str'):
            ent.append(i)                                # Adicionando Todas as EntidadeCod em uma lista
        desc = []                                        # Criando lista vazia
        for i in arq_codsusep['Empresa']:
            desc.append(i)                                # Adicionando Todos os Nomes de Empresa em uma lista

        mod = dict(zip(ent, desc))                        # Crando dicionario
        #Transformalo em Data Frame
        df_mod_e = pd.DataFrame(columns=['ENTCODIGO','Empresa'])
        df_mod_e.loc[:,'ENTCODIGO'] = mod.keys()
        df_mod_e.loc[:,'Empresa'] = mod.values()

        ent = []                                          # Criando lista vazia
        for i in arq_ramcod['ramo_s1'].astype('str'):
            ent.append(i)                                # Adicionando Todas as EntidadeCod em uma lista
        desc = []                                        # Criando lista vazia
        for i in arq_ramcod['ramNome']:
            desc.append(i)                                # Adicionando Todos os Nomes de Empresa em uma lista

        mod = dict(zip(ent, desc))                        # Crando dicionario
        #Transformalo em Data Frame
        df_mod_r = pd.DataFrame(columns=['RAMCODIGO','Desc_ramo'])
        df_mod_r.loc[:,'RAMCODIGO'] = mod.keys()
        df_mod_r.loc[:,'Desc_ramo'] = mod.values()

        # Juntando empresa com descrição a partir do arquivo modelo
        df_main = pd.merge(df_main,df_mod_e,on='ENTCODIGO',how='inner')
        df_main = pd.merge(df_main,df_mod_r,on='RAMCODIGO',how='inner')

        # Funcao para tornar os TPMOIDs negativos (Quando necessario)
        def valor_negativo_mov(column):
            TPMOID = column[1]
            fESRVALORMOV = column[0]
            
            if TPMOID == '0005':
                return -fESRVALORMOV
            else:
                return fESRVALORMOV
            
        # Aplicando funcao na coluna fESRVALORMOV
        df_main['fESRVALORMOV'] = df_main[['fESRVALORMOV','TPMOID']].apply(valor_negativo_mov,axis=1)

        # Funcao para tornar os TPMOIDs negativos (Quando necessario)
        def valor_negativo_mon(column):
            TPMOID = column[1]
            fESRVALORMON = column[0]
            
            if TPMOID == '0005':
                return -fESRVALORMON
            else:
                return fESRVALORMON
        
        # Aplicando funcao na coluna fESRVALORMON
        df_main['fESRVALORMON'] = df_main[['fESRVALORMON','TPMOID']].apply(valor_negativo_mon,axis=1)

        ################################################
        # Adicionando Nome do Criador do output e data #
        ################################################

        df_main.loc[:,'Criador'] = nome ### insere a coluna com o nome do criador da tabela
        df_main.loc[:,'Data do output'] = data_presente ### insere a coluna com a data de criacao da tabela

        # Arrumando Datas
        df_main['DTMRFMESANO']      = df_main['DTMRFMESANO'].apply(lambda x: x.strftime("%d/%m/%Y"))
        df_main['MES_ano']           = df_main['MES_ano'].apply(lambda x: x.strftime("%m/%Y"))
        df_main['MRFMES_ano']        = df_main['MRFMES_ano'].apply(lambda x: x.strftime("%d/%m/%Y"))  
        df_main['DTESRDATAINICIO']  = df_main['DTESRDATAINICIO'].apply(lambda x: x.strftime("%d/%m/%Y"))
        df_main['DTESRDATAFIM']     = df_main['DTESRDATAFIM'].apply(lambda x: x.strftime("%d/%m/%Y")) 
        df_main['DTESRDATAOCORR']   = df_main['DTESRDATAOCORR'].apply(lambda x: x.strftime("%d/%m/%Y"))
        df_main['DTESRDATAREG']     = df_main['DTESRDATAREG'].apply(lambda x: x.strftime("%d/%m/%Y"))
        df_main['DTESRDATACOMUNICA']     = df_main['DTESRDATACOMUNICA'].apply(lambda x: x.strftime("%d/%m/%Y"))

        # Criando DF_Criticas com valores em lista
        df_criticas = pd.DataFrame(criticas)
        df_criticas = df_criticas.rename(columns={0:'ID da Critica',1:'Descricao',2:'Codigo Invalido'})

        # Criando DF_Criticas2 com valores em dfs
        df_criticas_2 = pd.concat([df16_i,df16_f,df16_o,df16_r,df16_c,df17,df18,df19,df20,df21,df22,df23_1,df23_3,df23_4,df25,df26,df27,df29_p,df29_f])
        if len(df_criticas_2) == 0:
            df_criticas_2 = pd.DataFrame(columns = ['ID da Critica','RAMCODIGO','DTMRFMESANO','ENTCODIGO'])

        # Criando df filtro para as crticas_2
        df_filtado = pd.DataFrame(df_criticas_2[['ID da Critica','RAMCODIGO','DTMRFMESANO','ENTCODIGO']])
        df_filtado['Linha da Critica'] = df_filtado.index # Criando coluna com a linha da critica

        # Criando df geral e concatenando todas as criticas
        df_criticas_geral = pd.concat([df_criticas_i,df_filtado,df_criticas])
        df_criticas_geral.loc[:,'ENTCODIGO'] = itemtentcod  ### insere a coluna ID da Empresa com a da tabela

        # Juntando empresa com descrição a partir do arquivo modelo
        df_criticas_geral = pd.merge(df_criticas_geral,df_mod_e,on='ENTCODIGO',how='left')
        df_criticas_geral = pd.merge(df_criticas_geral,df_mod_r,on='RAMCODIGO',how='left')
        df_criticas_geral.loc[:,'Criador'] = nome ### insere a coluna com o nome do criador da tabela
        df_criticas_geral.loc[:,'Data do output'] = data_presente ### insere a coluna com a data de criacao da tabela
        df_criticas_geral = df_criticas_geral[['ID da Critica','Descricao','Linha da Critica','DTMRFMESANO','ENTCODIGO','Empresa','RAMCODIGO','Desc_ramo','Criador','Data do output']]

        return(df_main,df_criticas_geral)

        #######################
        ##### Outputs CSV #####
        #######################

    def outputs(df_main,df_criticas_geral):
        if len(df_criticas_geral) == 0:                             # Verificando se a lista de criticas esta vazia
            print('Nenhuma Critica encontrada no Quadro 376')
        else:
            print(df_criticas_geral['ID da Critica'].value_counts())
        df_criticas_geral.to_csv('C:\\Users\\haylton.neto\\OneDrive - GRANT THORNTON BRASIL\\Projetos_Digital_em_andamento\\Proj_Python_Versao1\\outputs\\376_Criticas_Consistencia.csv')         # Gerando csv das criticas impeditivas
        df_main.to_csv('C:\\Users\\haylton.neto\\OneDrive - GRANT THORNTON BRASIL\\Projetos_Digital_em_andamento\\Proj_Python_Versao1\\outputs\\376_Trabalho.csv')      # Gerando csv do Arquivo completo


        ############################################
        # Exportando arquivos finais para o SQLite #
        ############################################

        import sqlite3
        con = sqlite3.connect('C:\\Users\\haylton.neto\\OneDrive - GRANT THORNTON BRASIL\\Projetos_Digital_em_andamento\\Proj_Python_Versao1\\outputs\\DataBase.db')
        # Inserindo os registros do DataFrame df_criticas_geral
        df_main.to_sql(name='T_CONSISTENCIA', con=con)
        # Inserindo os registros do DataFrame df_criticas_geral
        df_criticas_geral.to_sql(name='T_CRITICAS', con=con)
        # Gravando a transação
        con.commit()
        # Fecha a conexão
        con.close()


