
#!/usr/bin/env python3
"""
🔧 INSTALADOR AUTOMÁTICO DE DEPENDÊNCIAS
=======================================

Script para instalar automaticamente todas as dependências necessárias
para o Gerador de Notas Fiscais.

Este script detecta automaticamente se você está usando ambiente virtual
e instala as dependências no local correto.
"""

import subprocess
import sys
import os
from pathlib import Path

def detectar_ambiente():
    """Detecta se está rodando em ambiente virtual"""
    # Verifica se VIRTUAL_ENV está definida
    if 'VIRTUAL_ENV' in os.environ:
        return True, os.environ['VIRTUAL_ENV']

    # Verifica se sys.prefix é diferente de sys.base_prefix (Python 3.3+)
    if hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix:
        return True, sys.prefix

    # Verifica se sys.real_prefix existe (Python 2/virtualenv antigo)
    if hasattr(sys, 'real_prefix'):
        return True, sys.prefix

    return False, None

def criar_ambiente_virtual():
    """Cria um ambiente virtual .venv se não existir"""
    caminho_venv = Path(".venv")

    if caminho_venv.exists():
        print(f"   ✅ Ambiente virtual já existe: {caminho_venv.absolute()}")
        return True, str(caminho_venv.absolute())

    print("🚀 CRIANDO AMBIENTE VIRTUAL...")
    print("=" * 40)
    print("   📍 Local: .venv")
    print("   ⏳ Isso pode levar alguns minutos...")
    print()

    try:
        # Comando para criar venv
        comando_criar = [sys.executable, "-m", "venv", ".venv"]

        resultado = subprocess.run(
            comando_criar,
            capture_output=True,
            text=True,
            timeout=180  # 3 minutos
        )

        if resultado.returncode == 0:
            print("   ✅ Ambiente virtual criado com sucesso!")

            # Verifica se foi criado corretamente
            if caminho_venv.exists():
                print(f"   📁 Localização: {caminho_venv.absolute()}")
                return True, str(caminho_venv.absolute())
            else:
                print("   ❌ Erro: Pasta .venv não foi criada")
                return False, None
        else:
            erro_msg = resultado.stderr.strip() if resultado.stderr else "Erro desconhecido"
            print(f"   ❌ Erro ao criar ambiente virtual: {erro_msg}")
            return False, None

    except subprocess.TimeoutExpired:
        print("   ⏰ Timeout na criação do ambiente virtual (>3min)")
        return False, None
    except Exception as e:
        print(f"   ❌ Erro inesperado: {str(e)}")
        return False, None

def ativar_e_usar_venv():
    """Instrui como ativar e usar o ambiente virtual criado"""
    print("\n🔧 COMO USAR O AMBIENTE VIRTUAL:")
    print("=" * 40)
    print("Para usar o ambiente virtual em futuras execuções:")
    print()

    if os.name == 'nt':  # Windows
        print("🖥️  WINDOWS:")
        print("   .venv\\Scripts\\activate")
        print("   python iniciar_rpa.py")
        script_path = Path(".venv/Scripts/python.exe")
    else:  # Unix/Mac
        print("🐧 LINUX/MAC:")
        print("   source .venv/bin/activate")
        print("   python iniciar_rpa.py")
        script_path = Path(".venv/bin/python")

    print()
    print("💡 OU execute diretamente:")
    if os.name == 'nt':
        print("   .venv\\Scripts\\python.exe iniciar_rpa.py")
    else:
        print("   .venv/bin/python iniciar_rpa.py")

    return str(script_path) if script_path.exists() else sys.executable

def main():
    print("🔧 INSTALADOR DE DEPENDÊNCIAS - GERADOR DE NOTAS FISCAIS")
    print("=" * 60)

    # Detecta ambiente
    em_venv, caminho_venv = detectar_ambiente()

    print("📍 DETECTANDO AMBIENTE:")
    if em_venv:
        print(f"   ✅ Ambiente virtual detectado: {caminho_venv}")
        print(f"   📦 Python: {sys.executable}")
        python_executavel = sys.executable
    else:
        print("   ⚠️  Usando Python do sistema")
        print(f"   📦 Python: {sys.executable}")
        print("   💡 Recomenda-se usar ambiente virtual (.venv)")

        # Pergunta se quer criar ambiente virtual
        print()
        criar_venv = input("🤔 Deseja criar um ambiente virtual (.venv) automaticamente? (S/N): ").strip().upper()

        if criar_venv == 'S':
            print()
            sucesso_venv, caminho_venv = criar_ambiente_virtual()

            if sucesso_venv:
                print("   🎉 Ambiente virtual criado com sucesso!")

                # Define o executável Python do venv para instalar as dependências
                if os.name == 'nt':  # Windows
                    python_executavel = str(Path(".venv/Scripts/python.exe"))
                else:  # Unix/Mac
                    python_executavel = str(Path(".venv/bin/python"))

                # Verifica se o executável existe
                if not Path(python_executavel).exists():
                    print(f"   ⚠️  Executável Python não encontrado em {python_executavel}")
                    print("   🔄 Usando Python do sistema para instalação...")
                    python_executavel = sys.executable
                else:
                    print(f"   📦 Usando Python do venv: {python_executavel}")

                em_venv = True  # Marca como em ambiente virtual para o resto do script
            else:
                print("   ❌ Falha na criação do ambiente virtual")
                print("   🔄 Continuando com Python do sistema...")
                python_executavel = sys.executable
        else:
            print("   ✅ Continuando com Python do sistema")
            python_executavel = sys.executable

    print()

    # Lista de dependências
    dependencias = ['pandas', 'selenium', 'openpyxl']

    print("📋 DEPENDÊNCIAS A INSTALAR:")
    for dep in dependencias:
        print(f"   • {dep}")
    print()

    # Confirma instalação
    resposta = input("🚀 Instalar todas as dependências? (S/N): ").strip().upper()
    if resposta != 'S':
        print("❌ Instalação cancelada.")
        return

    print("\n🔄 INICIANDO INSTALAÇÃO...")
    print("=" * 60)

    sucessos = []
    erros = []

    for dep in dependencias:
        print(f"\n📦 Instalando {dep}...")

        try:
            # Usa o Python definido (sistema ou venv) para instalar
            comando = [python_executavel, "-m", "pip", "install", dep]

            resultado = subprocess.run(
                comando,
                capture_output=True,
                text=True,
                timeout=300  # 5 minutos
            )

            if resultado.returncode == 0:
                print(f"   ✅ {dep} instalado com sucesso!")
                sucessos.append(dep)
            else:
                erro_msg = resultado.stderr.strip() if resultado.stderr else "Erro desconhecido"
                print(f"   ❌ Erro ao instalar {dep}")
                print(f"       {erro_msg}")
                erros.append((dep, erro_msg))

        except subprocess.TimeoutExpired:
            print(f"   ⏰ Timeout na instalação de {dep}")
            erros.append((dep, "Timeout (>5min)"))
        except Exception as e:
            print(f"   ❌ Erro inesperado: {str(e)}")
            erros.append((dep, str(e)))

    # Relatório final
    print("\n" + "=" * 60)
    print("📊 RELATÓRIO FINAL")
    print("=" * 60)

    if sucessos:
        print(f"✅ INSTALADAS COM SUCESSO ({len(sucessos)}):")
        for dep in sucessos:
            print(f"   • {dep}")

    if erros:
        print(f"\n❌ PROBLEMAS ({len(erros)}):")
        for dep, erro in erros:
            print(f"   • {dep}: {erro}")

    if not erros:
        print(f"\n🎉 PERFEITO! Todas as {len(dependencias)} dependências foram instaladas!")
        print("✅ Você pode agora executar o Gerador de Notas Fiscais")

        # Se criamos um venv, mostra instruções de uso
        if em_venv and Path(".venv").exists() and caminho_venv and ".venv" in str(caminho_venv):
            ativar_e_usar_venv()
    else:
        print(f"\n⚠️  {len(erros)} dependência(s) com problema.")
        print("\n💡 DICAS PARA RESOLVER:")
        print("1. Execute este script como ADMINISTRADOR")
        print("2. Atualize o pip: python -m pip install --upgrade pip")
        print("3. Tente instalar manualmente:")
        for dep, _ in erros:
            if python_executavel != sys.executable:
                print(f"   {python_executavel} -m pip install {dep}")
            else:
                print(f"   pip install {dep}")

    print("\n" + "=" * 60)
    input("Pressione ENTER para finalizar...")

if __name__ == "__main__":
    main()