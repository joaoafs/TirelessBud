import os
import sys
import ctypes
import subprocess
import time

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def clear_icon_cache():
    """
    Script para limpar o cache de ícones do Windows e forçar a atualização
    dos ícones do TirelessBud.
    """
    print("Iniciando limpeza do cache de ícones do Windows...")
    
    # Verifica se está rodando como administrador
    if not is_admin():
        print("Este script precisa ser executado como administrador.")
        print("Tentando reiniciar com privilégios elevados...")
        
        if sys.executable.endswith("pythonw.exe"):
            python_exe = sys.executable.replace("pythonw.exe", "python.exe")
        else:
            python_exe = sys.executable
            
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", python_exe, f'"{os.path.abspath(__file__)}"', None, 1
        )
        return
    
    # Caminhos para o cache de ícones
    icon_cache_paths = [
        os.path.join(os.environ['LOCALAPPDATA'], 'IconCache.db'),
        os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft\\Windows\\Explorer\\iconcache_*.db'),
        os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft\\Windows\\Explorer\\thumbcache_*.db')
    ]
    
    # Tenta parar o explorer
    try:
        print("Parando o Windows Explorer...")
        subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], 
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        time.sleep(2)
    except Exception as e:
        print(f"Erro ao parar o Explorer: {e}")
    
    # Remove arquivos de cache
    for cache_path in icon_cache_paths:
        try:
            if '*' in cache_path:
                # Se tiver wildcard, use dir para listar os arquivos
                import glob
                for file in glob.glob(cache_path):
                    try:
                        os.remove(file)
                        print(f"Removido: {file}")
                    except Exception as e:
                        print(f"Erro ao remover {file}: {e}")
            elif os.path.exists(cache_path):
                os.remove(cache_path)
                print(f"Removido: {cache_path}")
        except Exception as e:
            print(f"Erro ao processar {cache_path}: {e}")
    
    # Reinicia o explorer
    try:
        print("Reiniciando o Windows Explorer...")
        subprocess.Popen("explorer.exe")
    except Exception as e:
        print(f"Erro ao reiniciar o Explorer: {e}")
    
    print("\nLimpeza de cache concluída!")
    print("O ícone do TirelessBud deve ser atualizado agora.")
    print("\nSe o ícone ainda não estiver correto, tente:")
    print("1. Reiniciar o computador")
    print("2. Limpar o cache de miniaturas manualmente no Windows")
    print("3. Reconstruir o executável com build_exe_safe.py\n")
    
    input("Pressione Enter para sair...")

if __name__ == "__main__":
    clear_icon_cache()
