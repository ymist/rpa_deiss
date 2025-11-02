#!/usr/bin/env python3
"""
🤖 GERADOR AUTOMÁTICO DE NOTAS FISCAIS
====================================

Programa para preenchimento automático de notas fiscais eletrônicas.
Desenvolvido para veterinários e profissionais da saúde animal.

INSTRUÇÕES DE USO:
1. Coloque o arquivo Excel com os dados na mesma pasta deste programa
2. Execute este arquivo
3. Siga as instruções na tela

Autor: Assistente IA
Versão: 2.0 - Amigável para usuários
"""

import os
import sys
import pandas as pd
from pathlib import Path

def limpar_tela():
    """Limpa a tela do terminal"""
    os.system('cls' if os.name == 'nt' else 'clear')

def mostrar_banner():
    """Mostra o banner inicial do programa"""
    print("="*70)
    print("🤖 GERADOR AUTOMÁTICO DE NOTAS FISCAIS")
    print("="*70)
    print("✨ Versão Simplificada para Usuários")
    print("📋 Preenche formulários de notas fiscais automaticamente")
    print("🎯 Suporte para veterinários Kleiton e Katia")
    print("="*70)
    print()

def verificar_dependencias():
    """Verifica e instala automaticamente as bibliotecas necessárias"""
    print("🔍 Verificando dependências do sistema...")
    print("=" * 50)

    dependencias = {
        'pandas': 'pandas',
        'selenium': 'selenium',
        'openpyxl': 'openpyxl'
    }

    faltando = []
    instaladas = []

    # Primeiro, verifica quais estão faltando
    for lib_name, pip_name in dependencias.items():
        try:
            __import__(lib_name)
            print(f"   ✅ {lib_name} - OK")
            instaladas.append(lib_name)
        except ImportError:
            print(f"   ❌ {lib_name} - FALTANDO")
            faltando.append((lib_name, pip_name))

    if not faltando:
        print("\n✅ Todas as dependências estão instaladas!\n")
        return True

    # Se há dependências faltando, tenta instalar automaticamente
    print(f"\n🔧 INSTALAÇÃO AUTOMÁTICA DE {len(faltando)} DEPENDÊNCIA(S)")
    print("=" * 50)
    print("⏳ Aguarde, isso pode levar alguns minutos...")
    print()

    import subprocess
    import sys

    sucesso_instalacao = []
    erro_instalacao = []

    for lib_name, pip_name in faltando:
        print(f"📦 Instalando {lib_name}...")

        try:
            # Tenta instalar usando pip
            resultado = subprocess.run(
                [sys.executable, "-m", "pip", "install", pip_name],
                capture_output=True,
                text=True,
                timeout=300  # 5 minutos timeout
            )

            if resultado.returncode == 0:
                # Verifica se realmente foi instalado
                try:
                    __import__(lib_name)
                    print(f"   ✅ {lib_name} instalado com sucesso!")
                    sucesso_instalacao.append(lib_name)
                except ImportError:
                    print(f"   ❌ {lib_name} instalado mas não pode ser importado")
                    erro_instalacao.append((lib_name, "Erro na importação após instalação"))
            else:
                erro_msg = resultado.stderr.strip() if resultado.stderr else "Erro desconhecido"
                print(f"   ❌ Erro ao instalar {lib_name}: {erro_msg}")
                erro_instalacao.append((lib_name, erro_msg))

        except subprocess.TimeoutExpired:
            print(f"   ⏰ Timeout ao instalar {lib_name} (mais de 5 minutos)")
            erro_instalacao.append((lib_name, "Timeout na instalação"))
        except Exception as e:
            print(f"   ❌ Erro inesperado ao instalar {lib_name}: {str(e)}")
            erro_instalacao.append((lib_name, str(e)))

    # Relatório final
    print("\n" + "=" * 50)
    print("📊 RELATÓRIO DE INSTALAÇÃO")
    print("=" * 50)

    if sucesso_instalacao:
        print("✅ INSTALADAS COM SUCESSO:")
        for lib in sucesso_instalacao:
            print(f"   • {lib}")

    if erro_instalacao:
        print("\n❌ PROBLEMAS NA INSTALAÇÃO:")
        for lib, erro in erro_instalacao:
            print(f"   • {lib}: {erro}")

    # Se ainda há erros, oferece instalação manual
    if erro_instalacao:
        print(f"\n⚠️  {len(erro_instalacao)} dependência(s) não puderam ser instaladas automaticamente.")
        print("\n💡 SOLUÇÕES ALTERNATIVAS:")
        print("1. Execute este programa como ADMINISTRADOR")
        print("2. Instale manualmente com os comandos:")

        for lib_name, pip_name in faltando:
            if lib_name in [e[0] for e in erro_instalacao]:
                print(f"   pip install {pip_name}")

        print("3. Verifique se Python e pip estão atualizados")
        print("4. Se usa ambiente virtual (.venv), ative-o primeiro")

        continuar = input("\n❓ Tentar continuar mesmo assim? (S/N): ").strip().upper()
        if continuar != 'S':
            print("❌ Instalação cancelada. Resolva os problemas e tente novamente.")
            return False
        else:
            print("⚠️  Continuando... algumas funcionalidades podem não funcionar.")
            return True

    print(f"\n🎉 Todas as {len(dependencias)} dependências estão prontas!")
    print("=" * 50)
    return True

def encontrar_arquivo_excel():
    """Encontra automaticamente arquivos Excel na pasta"""
    pasta_atual = Path(".")
    arquivos_excel = list(pasta_atual.glob("*.xlsx")) + list(pasta_atual.glob("*.xls"))

    if not arquivos_excel:
        print("❌ ERRO: Nenhum arquivo Excel encontrado!")
        print("\n📝 INSTRUÇÕES:")
        print("1. Coloque seu arquivo Excel (.xlsx ou .xls) nesta pasta:")
        print(f"   {os.path.abspath('.')}")
        print("2. Execute este programa novamente")
        print()
        input("Pressione ENTER para sair...")
        return None

    if len(arquivos_excel) == 1:
        arquivo = arquivos_excel[0]
        print(f"📁 Arquivo Excel encontrado: {arquivo.name}")
        return str(arquivo)

    # Múltiplos arquivos - deixa usuário escolher
    print("📁 Múltiplos arquivos Excel encontrados:")
    for i, arquivo in enumerate(arquivos_excel, 1):
        print(f"   {i}. {arquivo.name}")

    while True:
        try:
            escolha = input(f"\nDigite o número do arquivo (1-{len(arquivos_excel)}): ").strip()
            indice = int(escolha) - 1
            if 0 <= indice < len(arquivos_excel):
                return str(arquivos_excel[indice])
            else:
                print("❌ Número inválido! Tente novamente.")
        except ValueError:
            print("❌ Digite apenas números!")

def validar_arquivo_excel(caminho_arquivo):
    """Valida se o arquivo Excel tem as colunas necessárias com validações detalhadas"""
    print(f"🔍 Validando arquivo: {os.path.basename(caminho_arquivo)}")
    print("=" * 60)

    try:
        # Lê o arquivo Excel
        df = pd.read_excel(caminho_arquivo)

        # Colunas obrigatórias e opcionais
        colunas_obrigatorias = ['Nome_Cliente', 'Nome_Pet', 'CPF', 'Valor']
        colunas_opcionais = ['Data', 'Cidade', 'Endereco']

        print(f"📊 INFORMAÇÕES GERAIS:")
        print(f"   • Total de registros: {len(df)}")
        print(f"   • Total de colunas: {len(df.columns)}")
        print(f"   • Colunas encontradas: {list(df.columns)}")
        print()

        # 1. VERIFICAÇÃO DE COLUNAS OBRIGATÓRIAS
        print("🔸 VERIFICAÇÃO DE COLUNAS:")
        faltando = []
        for coluna in colunas_obrigatorias:
            if coluna in df.columns:
                print(f"   ✅ {coluna} - OK")
            else:
                print(f"   ❌ {coluna} - FALTANDO")
                faltando.append(coluna)

        # Colunas opcionais
        for coluna in colunas_opcionais:
            if coluna in df.columns:
                print(f"   ✅ {coluna} - OK (opcional)")
            else:
                print(f"   ⚠️  {coluna} - Não encontrada (opcional)")

        if faltando:
            print(f"\n❌ ERRO CRÍTICO: Colunas obrigatórias faltando: {faltando}")
            print("\n📝 FORMATO CORRETO DO EXCEL:")
            print("   🔸 Nome_Cliente (texto) - Nome completo do cliente")
            print("   🔸 Nome_Pet (texto) - Nome do pet")
            print("   🔸 CPF (número) - CPF sem pontos e traços")
            print("   🔸 Valor (número) - Valor do serviço em reais")
            print("   🔸 Data (data) - Data do serviço DD/MM/AA (opcional)")
            print("   🔸 Cidade (texto) - Cidade do cliente (opcional)")
            print("   🔸 Endereco (texto) - Endereço completo (opcional)")
            return False, None

        # 2. VERIFICAÇÃO DE DADOS VAZIOS
        print("\n🔸 VERIFICAÇÃO DE DADOS VAZIOS:")
        registros_problema = 0
        for coluna in colunas_obrigatorias:
            vazios = df[coluna].isna().sum()
            if vazios > 0:
                print(f"   ⚠️  {coluna}: {vazios} registros vazios")
                registros_problema += vazios
            else:
                print(f"   ✅ {coluna}: todos preenchidos")

        # 3. VALIDAÇÃO ESPECÍFICA DOS DADOS
        print("\n🔸 VALIDAÇÃO DOS DADOS:")

        # Validar CPFs
        cpfs_invalidos = 0
        if 'CPF' in df.columns:
            for i, cpf in df['CPF'].items():
                if pd.notna(cpf):
                    # Limpeza similar à do RPA
                    cpf_str = str(cpf).replace('.0', '').replace('.', '').replace('-', '').replace(' ', '')
                    if not cpf_str.isdigit() or len(cpf_str) != 11:
                        cpfs_invalidos += 1

            if cpfs_invalidos > 0:
                print(f"   ⚠️  CPF: {cpfs_invalidos} CPFs inválidos (devem ter 11 dígitos)")
            else:
                print(f"   ✅ CPF: todos válidos")

        # Validar valores
        valores_invalidos = 0
        if 'Valor' in df.columns:
            for i, valor in df['Valor'].items():
                if pd.notna(valor):
                    try:
                        float(str(valor).replace(',', '.'))
                    except:
                        valores_invalidos += 1

            if valores_invalidos > 0:
                print(f"   ⚠️  Valor: {valores_invalidos} valores inválidos")
            else:
                print(f"   ✅ Valor: todos válidos")

        # Verifica se há dados
        if len(df) == 0:
            print("\n❌ ERRO: Arquivo Excel está vazio!")
            return False, None

        # 4. PREVIEW DETALHADO DOS DADOS
        print("\n" + "=" * 60)
        print("📋 PREVIEW DETALHADO DOS DADOS")
        print("=" * 60)

        preview_limit = min(5, len(df))
        for i in range(preview_limit):
            row = df.iloc[i]
            print(f"\n📌 REGISTRO {i+1}:")
            print(f"   👤 Cliente: {row.get('Nome_Cliente', 'N/A')}")
            print(f"   🐕 Pet: {row.get('Nome_Pet', 'N/A')}")

            # CPF formatado
            cpf_raw = row.get('CPF', 'N/A')
            if pd.notna(cpf_raw):
                # Limpeza igual à do RPA
                cpf_clean = str(cpf_raw).replace('.0', '').replace('.', '').replace('-', '').replace(' ', '')
                if len(cpf_clean) == 11 and cpf_clean.isdigit():
                    cpf_formatted = f"{cpf_clean[:3]}.{cpf_clean[3:6]}.{cpf_clean[6:9]}-{cpf_clean[9:]}"
                    print(f"   🆔 CPF: {cpf_formatted}")
                else:
                    print(f"   ⚠️  CPF: {cpf_raw} → {cpf_clean} (FORMATO INVÁLIDO)")
            else:
                print(f"   ❌ CPF: Não informado")

            # Valor formatado
            valor_raw = row.get('Valor', 'N/A')
            if pd.notna(valor_raw):
                try:
                    valor_num = float(str(valor_raw).replace(',', '.'))
                    print(f"   💰 Valor: R$ {valor_num:.2f}")
                except:
                    print(f"   ⚠️  Valor: {valor_raw} (FORMATO INVÁLIDO)")
            else:
                print(f"   ❌ Valor: Não informado")

            # Dados opcionais
            if 'Data' in df.columns and pd.notna(row.get('Data')):
                print(f"   📅 Data: {row.get('Data')}")

            if 'Cidade' in df.columns and pd.notna(row.get('Cidade')):
                print(f"   🏘️  Cidade: {row.get('Cidade')}")

            if 'Endereco' in df.columns and pd.notna(row.get('Endereco')):
                print(f"   📍 Endereço: {row.get('Endereco')}")

        if len(df) > preview_limit:
            print(f"\n... e mais {len(df) - preview_limit} registros")

        # 5. RESUMO DA VALIDAÇÃO
        print("\n" + "=" * 60)
        print("📊 RESUMO DA VALIDAÇÃO")
        print("=" * 60)

        total_problemas = registros_problema + cpfs_invalidos + valores_invalidos

        if total_problemas == 0:
            print("✅ ARQUIVO PERFEITO!")
            print("   • Todas as colunas obrigatórias presentes")
            print("   • Todos os dados válidos")
            print("   • Pronto para processamento")
        else:
            print(f"⚠️  ARQUIVO COM {total_problemas} PROBLEMA(S)")
            print("   • O programa tentará processar mesmo assim")
            print("   • Registros com problemas podem falhar")
            print("   • Recomenda-se corrigir o arquivo Excel")

        print(f"\n🎯 TOTAL DE NOTAS A PROCESSAR: {len(df)}")
        print("=" * 60)

        # Pergunta se quer continuar
        if total_problemas > 0:
            continuar = input("\n⚠️  Encontrados problemas nos dados. Continuar mesmo assim? (S/N): ").strip().upper()
            if continuar != 'S':
                print("❌ Operação cancelada. Corrija o arquivo Excel e tente novamente.")
                return False, None

        return True, df

    except Exception as e:
        print(f"❌ ERRO CRÍTICO ao ler arquivo Excel: {str(e)}")
        print("\n💡 POSSÍVEIS CAUSAS:")
        print("   • Arquivo corrompido ou não é Excel válido")
        print("   • Arquivo está aberto em outro programa (feche o Excel)")
        print("   • Problema de permissões de arquivo")
        print("   • Formato de arquivo não suportado")
        return False, None

def mostrar_estatisticas_detalhadas(df):
    """Mostra análise dos dados focada em problemas e validações"""
    print("\n" + "=" * 60)
    print("🔍 ANÁLISE DOS DADOS")
    print("=" * 60)

    # Resumo básico
    print(f"📊 RESUMO:")
    print(f"   • Total de registros: {len(df)}")
    print(f"   • Total de clientes únicos: {df['Nome_Cliente'].nunique() if 'Nome_Cliente' in df.columns else 'N/A'}")

    # Análise de problemas nos dados
    problemas_encontrados = []

    # Verificar dados vazios
    print(f"\n🔸 VERIFICAÇÃO DE DADOS VAZIOS:")
    colunas_obrigatorias = ['Nome_Cliente', 'Nome_Pet', 'CPF', 'Valor']
    for coluna in colunas_obrigatorias:
        if coluna in df.columns:
            vazios = df[coluna].isna().sum()
            if vazios > 0:
                print(f"   ⚠️  {coluna}: {vazios} registros vazios")
                problemas_encontrados.append(f"{vazios} registros sem {coluna}")
            else:
                print(f"   ✅ {coluna}: todos preenchidos")

    # Verificar CPFs inválidos
    print(f"\n🔸 VERIFICAÇÃO DE CPFs:")
    cpfs_problemas = 0
    if 'CPF' in df.columns:
        for i, cpf in df['CPF'].items():
            if pd.notna(cpf):
                cpf_clean = str(cpf).replace('.0', '').replace('.', '').replace('-', '').replace(' ', '')
                if not cpf_clean.isdigit() or len(cpf_clean) != 11:
                    cpfs_problemas += 1

        if cpfs_problemas > 0:
            print(f"   ⚠️  {cpfs_problemas} CPFs com formato inválido")
            problemas_encontrados.append(f"{cpfs_problemas} CPFs inválidos")
        else:
            print(f"   ✅ Todos os CPFs são válidos")

    # Verificar valores inválidos
    print(f"\n🔸 VERIFICAÇÃO DE VALORES:")
    valores_problemas = 0
    if 'Valor' in df.columns:
        for i, valor in df['Valor'].items():
            if pd.notna(valor):
                try:
                    float(str(valor).replace(',', '.'))
                except:
                    valores_problemas += 1

        if valores_problemas > 0:
            print(f"   ⚠️  {valores_problemas} valores com formato inválido")
            problemas_encontrados.append(f"{valores_problemas} valores inválidos")
        else:
            print(f"   ✅ Todos os valores são válidos")

    # Análise de cidades (apenas se houver dados opcionais problemáticos)
    if 'Cidade' in df.columns:
        cidades_vazias = df['Cidade'].isna().sum()
        if cidades_vazias > 0:
            print(f"\n🔸 DADOS OPCIONAIS:")
            print(f"   ⚠️  Cidade: {cidades_vazias} registros sem cidade (normal para alguns clientes)")

    # Resumo final dos problemas
    print("\n" + "=" * 60)
    if problemas_encontrados:
        print(f"⚠️  PROBLEMAS IDENTIFICADOS:")
        for problema in problemas_encontrados:
            print(f"   • {problema}")
        print(f"\n💡 O programa tentará processar mesmo assim")
    else:
        print("✅ NENHUM PROBLEMA ENCONTRADO!")
        print("   Todos os dados estão em formato válido")

    print("=" * 60)

def mostrar_preview_completo(df):
    """Mostra preview de todos os registros com paginação"""
    print("\n" + "=" * 60)
    print("📋 PREVIEW COMPLETO DE TODOS OS REGISTROS")
    print("=" * 60)

    registros_por_pagina = 10
    total_paginas = (len(df) + registros_por_pagina - 1) // registros_por_pagina
    pagina_atual = 1

    while True:
        inicio = (pagina_atual - 1) * registros_por_pagina
        fim = min(inicio + registros_por_pagina, len(df))

        print(f"\n📄 PÁGINA {pagina_atual} de {total_paginas} (Registros {inicio + 1} a {fim})")
        print("-" * 60)

        for i in range(inicio, fim):
            row = df.iloc[i]
            print(f"\n📌 REGISTRO {i + 1}:")
            print(f"   👤 Cliente: {row.get('Nome_Cliente', 'N/A')}")
            print(f"   🐕 Pet: {row.get('Nome_Pet', 'N/A')}")

            # CPF formatado
            cpf_raw = row.get('CPF', 'N/A')
            if pd.notna(cpf_raw):
                # Limpeza igual à do RPA
                cpf_clean = str(cpf_raw).replace('.0', '').replace('.', '').replace('-', '').replace(' ', '')
                if len(cpf_clean) == 11 and cpf_clean.isdigit():
                    cpf_formatted = f"{cpf_clean[:3]}.{cpf_clean[3:6]}.{cpf_clean[6:9]}-{cpf_clean[9:]}"
                    print(f"   🆔 CPF: {cpf_formatted}")
                else:
                    print(f"   ⚠️  CPF: {cpf_raw} → {cpf_clean} (INVÁLIDO)")
            else:
                print(f"   ❌ CPF: Não informado")

            # Valor formatado
            valor_raw = row.get('Valor', 'N/A')
            if pd.notna(valor_raw):
                try:
                    valor_num = float(str(valor_raw).replace(',', '.'))
                    print(f"   💰 Valor: R$ {valor_num:.2f}")
                except:
                    print(f"   ⚠️  Valor: {valor_raw} (INVÁLIDO)")
            else:
                print(f"   ❌ Valor: Não informado")

            # Dados opcionais
            if 'Data' in df.columns and pd.notna(row.get('Data')):
                print(f"   📅 Data: {row.get('Data')}")
            if 'Cidade' in df.columns and pd.notna(row.get('Cidade')):
                print(f"   🏘️  Cidade: {row.get('Cidade')}")
            if 'Endereco' in df.columns and pd.notna(row.get('Endereco')):
                print(f"   📍 Endereço: {row.get('Endereco')}")

        print("\n" + "-" * 60)

        if total_paginas == 1:
            input("Pressione ENTER para continuar...")
            break

        # Opções de navegação
        print("🔄 NAVEGAÇÃO:")
        opcoes = []
        if pagina_atual > 1:
            opcoes.append("A - Página anterior")
        if pagina_atual < total_paginas:
            opcoes.append("P - Próxima página")
        opcoes.append("C - Continuar para próximo passo")

        for opcao in opcoes:
            print(f"   {opcao}")

        escolha = input("\nSua escolha: ").strip().upper()

        if escolha == 'A' and pagina_atual > 1:
            pagina_atual -= 1
        elif escolha == 'P' and pagina_atual < total_paginas:
            pagina_atual += 1
        elif escolha == 'C':
            break
        else:
            if pagina_atual == 1:
                print("❌ Use 'P' para próxima ou 'C' para continuar")
            elif pagina_atual == total_paginas:
                print("❌ Use 'A' para anterior ou 'C' para continuar")
            else:
                print("❌ Use 'A', 'P' ou 'C'")

def escolher_cliente():
    """Interface amigável para escolher o cliente"""
    print("👥 SELEÇÃO DO CLIENTE")
    print("=" * 30)
    print("Escolha para qual cliente você quer gerar as notas:")
    print()
    print("1. 👨‍⚕️ Dr. KLEITON")
    print("   📍 Indaiatuba, SP")
    print("   🏥 Código atividade: 501")
    print("   💰 Alíquota: 2.01%")
    print()
    print("2. 👩‍⚕️ Dra. KATIA")
    print("   📍 Indaiatuba, SP")
    print("   🏥 Código atividade: 508")
    print("   💰 Alíquota: 2.01%")
    print()

    while True:
        escolha = input("Digite sua escolha (1 ou 2): ").strip()
        if escolha == '1':
            print("✅ Cliente selecionado: Dr. KLEITON")
            return 'kleiton'
        elif escolha == '2':
            print("✅ Cliente selecionado: Dra. KATIA")
            return 'katia'
        else:
            print("❌ Opção inválida! Digite 1 ou 2.")

def mostrar_instrucoes_navegador():
    """Mostra instruções para o navegador"""
    print("\n🌐 INSTRUÇÕES PARA O NAVEGADOR")
    print("=" * 40)
    print("📋 O navegador será aberto automaticamente.")
    print("💡 VOCÊ PRECISA FAZER:")
    print()
    print("1. 🔐 Fazer LOGIN no site da prefeitura")
    print("2. 📝 Navegar até a página de EMISSÃO DE NOTAS")
    print("3. ⚡ Deixar a página pronta para preenchimento")
    print("4. ⏳ Voltar aqui e pressionar ENTER")
    print()
    print("🚨 IMPORTANTE:")
    print("   • NÃO FECHE o navegador")
    print("   • NÃO MUDE de aba")
    print("   • Deixe o navegador visível na tela")
    print()

def confirmar_execucao(total_notas):
    """Confirma se o usuário quer executar"""
    print("\n🚀 PRONTO PARA EXECUTAR!")
    print("=" * 30)
    print(f"📊 Serão processadas {total_notas} notas fiscais")
    print("⏱️  Tempo estimado: aproximadamente {:.1f} minutos".format(total_notas * 0.5))
    print()
    print("🔄 MODO DE OPERAÇÃO:")
    print("   • TESTE: Apenas preenche os campos (recomendado)")
    print("   • PRODUÇÃO: Preenche E emite as notas")
    print()

    while True:
        modo = input("Escolha o modo (T para Teste, P para Produção): ").strip().upper()
        if modo == 'T':
            print("✅ Modo TESTE selecionado - apenas preenchimento")
            return True
        elif modo == 'P':
            confirmacao = input("⚠️  Modo PRODUÇÃO irá EMITIR as notas. Tem certeza? (S/N): ").strip().upper()
            if confirmacao == 'S':
                print("✅ Modo PRODUÇÃO confirmado")
                return False
            else:
                print("Voltando para seleção de modo...")
                continue
        else:
            print("❌ Digite T para Teste ou P para Produção")

def main():
    """Função principal simplificada"""
    try:
        limpar_tela()
        mostrar_banner()

        # 1. Verificar dependências
        if not verificar_dependencias():
            return

        # 2. Encontrar arquivo Excel
        caminho_excel = encontrar_arquivo_excel()
        if not caminho_excel:
            return

        # 3. Validar arquivo Excel e obter dados
        validacao_ok, df_dados = validar_arquivo_excel(caminho_excel)
        if not validacao_ok:
            input("Pressione ENTER para sair...")
            return

        # 4. Mostrar estatísticas detalhadas
        mostrar_estatisticas_detalhadas(df_dados)

        # Pergunta se quer ver preview completo
        if len(df_dados) > 5:
            ver_todos = input(f"\n🔍 Quer ver o preview de TODOS os {len(df_dados)} registros? (S/N): ").strip().upper()
            if ver_todos == 'S':
                mostrar_preview_completo(df_dados)

        # 5. Escolher cliente
        cliente = escolher_cliente()

        # 6. Usar dados já carregados da validação
        total_notas = len(df_dados)

        # 7. Confirmar execução
        modo_teste = confirmar_execucao(total_notas)

        # 8. Mostrar instruções do navegador
        mostrar_instrucoes_navegador()
        input("✅ Pressione ENTER quando estiver pronto...")

        # 9. Importar e executar o RPA
        print("\n🤖 Iniciando o robô...")
        print("=" * 30)

        try:
            from rpa_notas_fiscais import RPANotasFiscais

            # URL do site (pode ser configurada)
            url_site = "https://deiss.indaiatuba.sp.gov.br/Deiss/restrito/nf_emissao.jsf"

            # Criar e executar RPA (com delay otimizado)
            rpa = RPANotasFiscais(url_site, caminho_excel, cliente, delay=0.5)
            rpa.processar_notas(modo_teste=modo_teste)

        except ImportError:
            print("❌ ERRO: Arquivo 'rpa_notas_fiscais.py' não encontrado!")
            print("Certifique-se que ambos os arquivos estão na mesma pasta.")
        except Exception as e:
            print(f"❌ ERRO durante execução: {str(e)}")

    except KeyboardInterrupt:
        print("\n\n⏹️  Operação cancelada pelo usuário.")
    except Exception as e:
        print(f"\n❌ ERRO INESPERADO: {str(e)}")
        print("Entre em contato com o suporte técnico.")

    input("\nPressione ENTER para sair...")

if __name__ == "__main__":
    main()