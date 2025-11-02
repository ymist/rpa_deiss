#!/usr/bin/env python3
"""
Script para criar arquivo Excel de exemplo para o RPA de Notas Fiscais
Execute este script para gerar o arquivo notas_fiscais.xlsx com dados de teste
"""

import pandas as pd
from datetime import datetime

def criar_excel_exemplo():
    """Cria arquivo Excel de exemplo com dados fictícios"""
    
    # Dados de exemplo com CPFs e informações fictícias
    dados = [
        {
            'Data': '16/09/25',
            'Nome_Cliente': 'Maria Silva Santos',
            'CPF': '11122233344',
            'Nome_Item': 'Serviço A',
            'Tipo_Servico': 'Consulta Padrão',
            'Valor': 350.00,
            'Endereco': 'R das Flores, 164',
            'Cidade': 'CIDADE_A'
        },
        {
            'Data': '16/09/25',
            'Nome_Cliente': 'João Pedro Costa',
            'CPF': '22233344455',
            'Nome_Item': 'Serviço B',
            'Tipo_Servico': 'Atendimento Especial',
            'Valor': 300.00,
            'Endereco': 'AV Principal, 1335',
            'Cidade': 'CIDADE_B'
        },
        {
            'Data': '17/09/25',
            'Nome_Cliente': 'Ana Paula Oliveira',
            'CPF': '33344455566',
            'Nome_Item': 'Serviço C',
            'Tipo_Servico': 'Consulta Técnica',
            'Valor': 350.00,
            'Endereco': 'R 24 de Maio, 338',
            'Cidade': 'CIDADE_C'
        },
        {
            'Data': '18/09/25',
            'Nome_Cliente': 'Carlos Eduardo Lima',
            'CPF': '44455566677',
            'Nome_Item': 'Serviço D',
            'Tipo_Servico': 'Manutenção',
            'Valor': 300.00,
            'Endereco': 'R das Palmeiras, 143',
            'Cidade': 'CIDADE_D'
        },
        {
            'Data': '19/09/25',
            'Nome_Cliente': 'Patricia Fernandes',
            'CPF': '55566677788',
            'Nome_Item': 'Serviço E',
            'Tipo_Servico': 'Instalação',
            'Valor': 300.00,
            'Endereco': 'R do Comércio, 100',
            'Cidade': 'CIDADE_A'
        },
        {
            'Data': '19/05/25',
            'Nome_Cliente': 'Roberto Alves',
            'CPF': '66677788899',
            'Nome_Item': 'Serviço F',
            'Tipo_Servico': 'Configuração',
            'Valor': 330.00,
            'Endereco': 'AV Central, 181',
            'Cidade': 'CIDADE_B'
        },
        {
            'Data': '20/09/25',
            'Nome_Cliente': 'Juliana Martins',
            'CPF': '77788899900',
            'Nome_Item': 'Serviço G',
            'Tipo_Servico': 'Atualização',
            'Valor': 150.00,
            'Endereco': 'R das Acácias, 456',
            'Cidade': 'CIDADE_EXEMPLO'
        },
        {
            'Data': '21/09/25',
            'Nome_Cliente': 'Fernando Santos',
            'CPF': '88899900011',
            'Nome_Item': 'Serviço H',
            'Tipo_Servico': 'Reparação',
            'Valor': 800.00,
            'Endereco': 'AV dos Estados, 789',
            'Cidade': 'CIDADE_C'
        },
        {
            'Data': '22/09/25',
            'Nome_Cliente': 'Amanda Costa',
            'CPF': '99900011122',
            'Nome_Item': 'Serviço I',
            'Tipo_Servico': 'Avaliação',
            'Valor': 120.00,
            'Endereco': 'R do Centro, 25',
            'Cidade': 'CIDADE_D'
        },
        {
            'Data': '23/09/25',
            'Nome_Cliente': 'Ricardo Oliveira',
            'CPF': '10011122233',
            'Nome_Item': 'Serviço J',
            'Tipo_Servico': 'Inspeção',
            'Valor': 250.00,
            'Endereco': 'R Nova Esperança, 88',
            'Cidade': 'CIDADE_A'
        }
    ]
    
    # Criar DataFrame
    df = pd.DataFrame(dados)
    
    # Configurar o ExcelWriter com formatação
    with pd.ExcelWriter('notas_fiscais.xlsx', engine='openpyxl') as writer:
        # Escrever dados na planilha
        df.to_excel(writer, sheet_name='Dados', index=False)
        
        # Obter a worksheet para formatação
        worksheet = writer.sheets['Dados']
        
        # Ajustar largura das colunas
        column_widths = {
            'A': 12,  # Data
            'B': 20,  # Nome_Cliente
            'C': 15,  # CPF
            'D': 15,  # Nome_Item
            'E': 20,  # Tipo_Servico
            'F': 10,  # Valor
            'G': 35,  # Endereco
            'H': 15   # Cidade
        }
        
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
        
        # Formatar cabeçalho
        for cell in worksheet[1]:
            cell.font = cell.font.copy(bold=True)
            cell.fill = cell.fill.copy(fgColor="DDDDDD")
    
    print("✅ Arquivo 'notas_fiscais.xlsx' criado com sucesso!")
    print(f"📊 Total de registros: {len(dados)}")
    print("\n📋 Estrutura criada:")
    print(df.head())
    print(f"\n📍 Cidades incluídas: {', '.join(df['Cidade'].unique())}")
    print(f"💰 Valores: R$ {df['Valor'].min():.2f} - R$ {df['Valor'].max():.2f}")

def validar_dados():
    """Valida o arquivo criado"""
    try:
        df = pd.read_excel('notas_fiscais.xlsx')
        
        print("\n🔍 Validação dos dados:")
        
        # Verificar colunas obrigatórias
        colunas_obrigatorias = ['Data', 'Nome_Cliente', 'CPF', 'Nome_Item', 'Tipo_Servico', 'Valor', 'Endereco', 'Cidade']
        colunas_faltando = [col for col in colunas_obrigatorias if col not in df.columns]
        
        if colunas_faltando:
            print(f"❌ Colunas faltando: {colunas_faltando}")
        else:
            print("✅ Todas as colunas obrigatórias estão presentes")
        
        # Verificar CPFs
        cpfs_invalidos = df[df['CPF'].astype(str).str.len() != 11]
        if len(cpfs_invalidos) > 0:
            print(f"⚠️  CPFs com formato incorreto: {len(cpfs_invalidos)}")
        else:
            print("✅ Todos os CPFs têm 11 dígitos")
        
        # Verificar valores
        valores_invalidos = df[df['Valor'] <= 0]
        if len(valores_invalidos) > 0:
            print(f"⚠️  Valores inválidos: {len(valores_invalidos)}")
        else:
            print("✅ Todos os valores são positivos")
        
        # Mostrar estatísticas
        print(f"\n📈 Estatísticas:")
        print(f"   • Total de registros: {len(df)}")
        print(f"   • Valor total: R$ {df['Valor'].sum():.2f}")
        print(f"   • Valor médio: R$ {df['Valor'].mean():.2f}")
        print(f"   • Cidades únicas: {df['Cidade'].nunique()}")
        print(f"   • Clientes únicos: {df['Nome_Cliente'].nunique()}")
        
        return True
        
    except FileNotFoundError:
        print("❌ Arquivo não encontrado. Execute primeiro a criação do arquivo.")
        return False
    except Exception as e:
        print(f"❌ Erro na validação: {str(e)}")
        return False

def criar_template_vazio():
    """Cria um template vazio para preenchimento manual"""
    
    # Criar DataFrame vazio com apenas os cabeçalhos
    template = pd.DataFrame(columns=[
        'Data', 'Nome_Cliente', 'CPF', 'Nome_Item', 
        'Tipo_Servico', 'Valor', 'Endereco', 'Cidade'
    ])
    
    # Salvar template
    with pd.ExcelWriter('template_notas_fiscais.xlsx', engine='openpyxl') as writer:
        template.to_excel(writer, sheet_name='Template', index=False)
        
        worksheet = writer.sheets['Template']
        
        # Ajustar larguras
        column_widths = {'A': 12, 'B': 20, 'C': 15, 'D': 15, 'E': 20, 'F': 10, 'G': 35, 'H': 15}
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
        
        # Formatar cabeçalho
        for cell in worksheet[1]:
            cell.font = cell.font.copy(bold=True)
            cell.fill = cell.fill.copy(fgColor="DDDDDD")
    
    print("✅ Template vazio criado: 'template_notas_fiscais.xlsx'")
    print("\n📝 Instruções de preenchimento:")
    print("   • Data: formato DD/MM/AA (ex: 16/09/25)")
    print("   • CPF: 11 dígitos sem pontos ou hífens (ex: 12345678901)")
    print("   • Valor: use ponto decimal (ex: 350.00)")
    print("   • Endereço: R Nome da Rua, Número (ex: R das Flores, 123)")

if __name__ == "__main__":
    print("📊 Gerador de Excel para RPA de Notas Fiscais")
    print("=" * 55)
    
    # Menu de opções
    print("\nEscolha uma opção:")
    print("1. Criar arquivo com dados de exemplo")
    print("2. Criar template vazio para preenchimento manual")
    print("3. Validar arquivo existente")
    print("4. Criar ambos (exemplo + template)")
    
    opcao = input("\nDigite sua opção (1-4): ").strip()
    
    if opcao == "1":
        criar_excel_exemplo()
        validar_dados()
    elif opcao == "2":
        criar_template_vazio()
    elif opcao == "3":
        validar_dados()
    elif opcao == "4":
        criar_excel_exemplo()
        criar_template_vazio()
        validar_dados()
    else:
        print("❌ Opção inválida. Execute novamente.")
    
    print("\n" + "=" * 55)
    print("✅ Arquivos prontos para usar com o RPA!")