import pandas as pd
from unidecode import unidecode


class Venda:
    def __init__(self, archive, index):
        self.data = archive['Data de venda'][index]
        self.cliente = Cliente(archive, index)
        self.produto = archive['Título do anúncio'][index]
        self.quantidade = archive['Unidades'][index]
        self.valor_total = archive['Total (BRL)'][index]
        self.regiao = 'ML'


class Cliente:
    def __init__(self, archive, index):
        self.nome_completo = 'NOME COMPLETO'
        self.primeiro_nome = archive['Nome do comprador'][index]
        self.segundo_nome = archive['Sobrenome do comprador'][index]
        self.cpf = archive['CPF'][index]
        self.cnpj = ' '
        self.ddd = '00'
        self.telefone = '99999999'
        self.endereco = Endereco(archive, index)
        self.juntar_nomes()
        self.cpf_or_cnpj()

    def juntar_nomes(self):
        self.nome_completo = str(self.primeiro_nome) + ' ' + str(self.segundo_nome)

    def cpf_or_cnpj(self):
        if len(str(self.cpf)) != 11:
            self.cnpj = self.cpf
            self.cpf = ' '


class Endereco:
    def __init__(self, archive, index):
        self.endereco_bruto = str(archive['Endereço'][index])
        self.rua = ' '
        self.numero = ' '
        self.complemento = ' '
        self.bairro = ' '
        self.cidade = archive['Cidade'][index]
        self.estado = archive['Estado'][index]
        self.cep = archive['CEP'][index]
        self.separar_endereco()
        self.estado_2_uf()

    def separar_endereco(self):
        if '/' in self.endereco_bruto:
            endereco_bruto = self.endereco_bruto
            rua = endereco_bruto.split('/')[0]
            rua = rua.split()
            self.numero = rua.pop()
            self.rua = rua[0]
            for j in range(len(rua) - 1):
                self.rua = self.rua + ' ' + rua[j+1]
            bairro = self.endereco_bruto.split(',')
            complemento = bairro
            if len(bairro) == 3:
                bairro = bairro[0].split(' - ')
                self.bairro = bairro[len(bairro)-1]
            complemento = complemento[0].split(' - ')
            if len(complemento) == 3:
                self.complemento = complemento[0].split('/').pop()
        else:
            self.cidade = ' '
            self.estado = ' '
            self.cep = ' '

    def estado_2_uf(self):
        self.estado = str(self.estado)
        if self.estado != ' ':
            lista_estados = ['Acre', 'Alagoas', 'Amapá', 'Amazonas', 'Bahia', 'Ceará', 'Distrito Federal',
                             'Espirito Santo', 'Goiás', 'Maranhão', 'Mato Grosso', 'Mato Grosso do Sul',
                             'Minas Gerais', 'Pará', 'Paraíba', 'Paraná', 'Pernambuco', 'Piauí', 'Rio de Janeiro',
                             'Rio Grande do Norte', 'Rio Grande do Sul', 'Rondônia', 'Roraima', 'Santa Catarina',
                             'São Paulo', 'Sergipe', 'Tocantins']
            lista_ufs = ['AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG', 'PA', 'PB',
                         'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO']
            for x in range(len(lista_estados)):
                if self.estado == lista_estados[x]:
                    self.estado = lista_ufs[x]


file = pd.read_excel('vendas.xlsx', converters={'CPF': str, 'CEP': str, 'Endereço': str}, skiprows=1)
vendas = []

for i in range(len(file['Nome do comprador'])):
    if not file['Nome do comprador'].isnull()[i]:
        venda = Venda(file, i)
        vendas.append(venda)
        print(' ')
        print('Data: ', venda.data)
        print('Nome: ', venda.cliente.nome_completo)
        print('CPF: ', venda.cliente.cpf)
        print('CNPJ: ', venda.cliente.cnpj)
        print('DDD: ', venda.cliente.ddd)
        print('Telefone: ', venda.cliente.telefone)
        print('Rua: ', venda.cliente.endereco.rua)
        print('Numero: ', venda.cliente.endereco.numero)
        print('Complemento: ', venda.cliente.endereco.complemento)
        print('Bairro: ', venda.cliente.endereco.bairro)
        print('Cidade: ', venda.cliente.endereco.cidade)
        print('Estado: ', venda.cliente.endereco.estado)
        print('CEP: ', venda.cliente.endereco.cep)
        print('Produto: ', venda.produto)
        print('Quantidade: ', venda.quantidade)
        print('Valor Total: ', venda.valor_total)
        print('Região: ', venda.regiao)

data = []
nome = []
cpf = []
cnpj = []
ddd = []
telefone = []
rua = []
numero = []
complemento = []
bairro = []
cidade = []
estado = []
cep = []
produto = []
quantidade = []
valor = []
regiao = []

for venda in vendas:
    data.append(venda.data)
    nome.append(unidecode(str(venda.cliente.nome_completo).upper()))
    cpf.append(venda.cliente.cpf)
    cnpj.append(venda.cliente.cnpj)
    ddd.append(venda.cliente.ddd)
    telefone.append(venda.cliente.telefone)
    rua.append(unidecode(str(venda.cliente.endereco.rua).upper()))
    numero.append(venda.cliente.endereco.numero)
    complemento.append(unidecode(str(venda.cliente.endereco.complemento).upper()))
    bairro.append(unidecode(str(venda.cliente.endereco.bairro).upper()))
    cidade.append(unidecode(str(venda.cliente.endereco.cidade).upper()))
    estado.append(unidecode(str(venda.cliente.endereco.estado).upper()))
    cep.append(venda.cliente.endereco.cep)
    produto.append(unidecode(str(venda.produto).upper()))
    quantidade.append(venda.quantidade)
    valor.append(venda.valor_total)
    regiao.append(venda.regiao)

dados = {
    'DATA': data,
    'NOME': nome,
    'CPF': cpf,
    'DDD': ddd,
    'TELEFONE': telefone,
    'ENDERECO': rua,
    'NUMERO': numero,
    'COMPLEMENTO': complemento,
    'BAIRRO': bairro,
    'CIDADE': cidade,
    'ESTADO': estado,
    'CEP': cep,
    'PRODUTO': produto,
    'QUANTIDADE': quantidade,
    'VALOR': valor,
    'REGIAO': regiao
}

output = pd.DataFrame(dados)

output.to_excel('cadastrar_clientes.xlsx', sheet_name='Cadastrar', index=None)