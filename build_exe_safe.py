import os
import subprocess
import sys
import time
import shutil
import re
import tempfile
import ctypes
import uuid
import platform

# Verificar se está rodando como administrador
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Função para extrair a versão do arquivo Python
def extract_version_from_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            # Busca por VERSION = "X.Y.Z" no arquivo
            version_match = re.search(r'VERSION\s*=\s*["\']([0-9]+\.[0-9]+\.[0-9]+)["\']', content)
            if version_match:
                return version_match.group(1)
    except Exception as e:
        print(f"Erro ao extrair versão do arquivo: {e}")
    
    # Retorna uma versão padrão se não conseguir extrair
    return "1.0.0"

# Função para criar arquivo de manifesto para UAC e DPI awareness
def create_manifest_file(build_dir, product_name):
    manifest_file = os.path.join(build_dir, "app.manifest")
    with open(manifest_file, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        f.write('<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">\n')
        f.write('  <assemblyIdentity version="1.0.0.0" name="' + product_name + '"/>\n')
        f.write('  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">\n')
        f.write('    <security>\n')
        f.write('      <requestedPrivileges>\n')
        f.write('        <requestedExecutionLevel level="asInvoker" uiAccess="false"/>\n')
        f.write('      </requestedPrivileges>\n')
        f.write('    </security>\n')
        f.write('  </trustInfo>\n')
        f.write('  <compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">\n')
        f.write('    <application>\n')
        f.write('      <!-- Windows 10 and 11 -->\n')
        f.write('      <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>\n')
        f.write('      <!-- Windows 8.1 -->\n')
        f.write('      <supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}"/>\n')
        f.write('      <!-- Windows 8 -->\n')
        f.write('      <supportedOS Id="{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"/>\n')
        f.write('      <!-- Windows 7 -->\n')
        f.write('      <supportedOS Id="{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"/>\n')
        f.write('    </application>\n')
        f.write('  </compatibility>\n')
        f.write('  <application xmlns="urn:schemas-microsoft-com:asm.v3">\n')
        f.write('    <windowsSettings>\n')
        f.write('      <dpiAware xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">true</dpiAware>\n')
        f.write('      <dpiAwareness xmlns="http://schemas.microsoft.com/SMI/2016/WindowsSettings">PerMonitorV2</dpiAwareness>\n')
        f.write('    </windowsSettings>\n')
        f.write('  </application>\n')
        f.write('</assembly>\n')
    
    return manifest_file

# Função para criar arquivo de metadados de versão para o executável
def create_version_file(build_dir, version_info):
    version_file = os.path.join(build_dir, "file_version_info.txt")
    with open(version_file, 'w', encoding='utf-8') as f:
        f.write('# UTF-8\n')
        f.write('#\n')
        f.write('# For more details about fixed file info \'ffi\' see:\n')
        f.write('# http://msdn.microsoft.com/en-us/library/ms646997.aspx\n')
        f.write('VSVersionInfo(\n')
        f.write('  ffi=FixedFileInfo(\n')
        f.write('    # filevers and prodvers should be always a tuple with four items: (1, 2, 3, 4)\n')
        f.write('    # Set not needed items to zero 0.\n')
        
        version_parts = version_info["FileVersion"].split('.')
        while len(version_parts) < 4:
            version_parts.append('0')
        version_tuple = ', '.join(version_parts)
        
        f.write(f'    filevers=({version_tuple}),\n')
        f.write(f'    prodvers=({version_tuple}),\n')
        f.write('    # Contains a bitmask that specifies the valid bits \'flags\'\n')
        f.write('    mask=0x3f,\n')
        f.write('    # Contains a bitmask that specifies the Boolean attributes of the file.\n')
        f.write('    flags=0x0,\n')
        f.write('    # The operating system for which this file was designed.\n')
        f.write('    # 0x4 - NT and there is no need to change it.\n')
        f.write('    OS=0x4,\n')
        f.write('    # The general type of file.\n')
        f.write('    # 0x1 - the file is an application.\n')
        f.write('    fileType=0x1,\n')
        f.write('    # The function of the file.\n')
        f.write('    # 0x0 - the function is not defined for this fileType\n')
        f.write('    subtype=0x0,\n')
        f.write('    # Creation date and time stamp.\n')
        f.write('    date=(0, 0)\n')
        f.write('    ),\n')
        f.write('  kids=[\n')
        f.write('    StringFileInfo(\n')
        f.write('      [\n')
        f.write('      StringTable(\n')
        f.write('        u\'040904B0\',\n')
        f.write('        [StringStruct(u\'CompanyName\', u\'' + version_info["CompanyName"] + '\'),\n')
        f.write('        StringStruct(u\'FileDescription\', u\'' + version_info["FileDescription"] + '\'),\n')
        f.write('        StringStruct(u\'FileVersion\', u\'' + version_info["FileVersion"] + '\'),\n')
        f.write('        StringStruct(u\'InternalName\', u\'' + version_info["InternalName"] + '\'),\n')
        f.write('        StringStruct(u\'LegalCopyright\', u\'' + version_info["LegalCopyright"] + '\'),\n')
        f.write('        StringStruct(u\'OriginalFilename\', u\'' + version_info["OriginalFilename"] + '\'),\n')
        f.write('        StringStruct(u\'ProductName\', u\'' + version_info["ProductName"] + '\'),\n')
        f.write('        StringStruct(u\'ProductVersion\', u\'' + version_info["ProductVersion"] + '\')])\n')
        f.write('      ]),\n')
        f.write('    VarFileInfo([VarStruct(u\'Translation\', [1033, 1200])])\n')
        f.write('  ]\n')
        f.write(')\n')
    
    return version_file

def build_exe_safe():
    """
    Constrói uma versão segura do executável TirelessBud usando PyInstaller
    com otimizações para evitar falsos positivos de antivírus.
    """
    print("Iniciando construção da versão segura otimizada do TirelessBud.exe...")
    print("Esta versão combina características de segurança com otimizações anti-falso-positivo.")
    
    # Caminho para o código fonte e ícone
    base_dir = os.path.dirname(os.path.abspath(__file__))
    source_path = os.path.join(base_dir, "code", "main_exe_safe.py")
    icon_path = os.path.join(base_dir, "Logo_TBud.ico")
    
    # Verificar e remover o ícone antigo se existir
    old_icon_path = os.path.join(base_dir, "LogoTBud.ico")
    if os.path.exists(old_icon_path):
        try:
            os.remove(old_icon_path)
            print(f"Ícone antigo removido: {old_icon_path}")
        except Exception as e:
            print(f"Erro ao remover ícone antigo: {e}")
    
    # Verificar se o novo ícone existe
    if not os.path.exists(icon_path):
        print(f"ERRO: Ícone novo não encontrado em {icon_path}")
        return False
    else:
        print(f"Usando ícone: {icon_path}")
    
    # Extrai a versão do arquivo fonte
    version = extract_version_from_file(source_path)
    exe_name = f"TirelessBud_v{version}"  # Removido "Safe" do nome do executável
    print(f"Versão detectada: {version}")
    print(f"Nome do executável: {exe_name}.exe")
    
    # Usar a pasta build no diretório do projeto
    build_dir = os.path.join(base_dir, "build")
    os.makedirs(build_dir, exist_ok=True)
    print(f"Usando diretório de build: {build_dir}")
    
    # Verificação de arquivos fonte
    if not os.path.exists(source_path):
        print(f"ERRO: Arquivo fonte seguro não encontrado.")
        print("Copiando o arquivo main.py para main_exe_safe.py...")
        try:
            shutil.copy2(
                os.path.join(base_dir, "code", "main.py"),
                os.path.join(base_dir, "code", "main_exe_safe.py")
            )
            print("Arquivo copiado. Por favor, edite code/main_exe_safe.py conforme instruções e execute este script novamente.")
            return False
        except Exception as e:
            print(f"Erro ao copiar arquivo: {e}")
            return False
    
    # Criar arquivos de metadados para o executável
    print("Gerando metadados para o executável...")
    company_name = "TirelessBud Software"
    product_name = "TirelessBud"
    
    version_info = {
        "CompanyName": company_name,
        "FileDescription": "TirelessBud - Ferramenta de Produtividade",  # Removido "Safe"
        "FileVersion": version,
        "InternalName": "TirelessBud",
        "LegalCopyright": f"© {time.strftime('%Y')} github.com/joaoafs",
        "OriginalFilename": f"{exe_name}.exe",
        "ProductName": product_name,
        "ProductVersion": version
    }
    
    # Criar arquivos de metadados
    version_file = create_version_file(build_dir, version_info)
    manifest_file = create_manifest_file(build_dir, product_name)
    
    # Definir estágios do PyInstaller para monitoramento
    build_stages = {
        r'INFO: PyInstaller:': 1,
        r'INFO: checking Analysis': 5,
        r'INFO: Building Analysis': 10,
        r'INFO: Analyzing ': 15,
        r'INFO: Processing module hooks': 25,
        r'INFO: Looking for ctypes DLLs': 35,
        r'INFO: Analyzing run-time hooks': 40,
        r'INFO: checking PYZ': 45,
        r'INFO: Building PYZ': 50,
        r'INFO: checking PKG': 60,
        r'INFO: Building PKG': 70,
        r'INFO: Bootloader': 80,
        r'INFO: checking EXE': 85,
        r'INFO: Building EXE': 90,
        r'INFO: Appending PKG': 95,
        r'INFO: Building EXE .* completed': 98
    }
    
    # Comando PyInstaller para usar a pasta build
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--clean",
        "--noconfirm",
        f"--name={exe_name}",
        f"--distpath={os.path.join(base_dir, 'dist')}",  # Salva diretamente na pasta dist
        f"--workpath={os.path.join(build_dir, 'work')}",  # Subdiretório work na pasta build
        f"--specpath={build_dir}",  # Coloca o arquivo .spec na pasta build
        "--noupx",  # Evita compressão UPX que causa mais falsos positivos
        "--hidden-import=pymupdf",
        "--hidden-import=PyPDF2",
        "--hidden-import=openpyxl",
        "--hidden-import=PIL",
        f"--add-data={icon_path};.",  # Use o caminho completo para o ícone
    ]
    
    # Adicione opcionalmente o manifesto e arquivo de versão se estiverem disponíveis
    if os.path.exists(icon_path):
        cmd.append(f"--icon={icon_path}")
    
    if version_file:
        cmd.append(f"--version-file={version_file}")
    
    if manifest_file:
        cmd.append(f"--manifest={manifest_file}")
    
    # Adiciona o arquivo fonte ao final
    cmd.append(source_path)
    
    print("Executando comando:", " ".join(cmd))
    print("Isto pode levar vários minutos...")
    print("\nAcompanhamento de progresso em tempo real:")
    print("-----------------------------------------")
    
    start_time = time.time()
    last_update_time = start_time
    
    # Função para desenhar barra de progresso
    def draw_progress_bar(progress, width=50):
        completed = int(width * progress / 100)
        remaining = width - completed
        bar = '█' * completed + '▒' * remaining
        return f"[{bar}] {progress}%"

    # Iniciar processo com captura de saída em tempo real
    process = subprocess.Popen(
        cmd, 
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        universal_newlines=True
    )
    
    # Variáveis para monitorar progresso
    current_progress = 0
    current_stage = "Iniciando"
    last_line = ""
    
    # Atualiza o caminho do executável para usar o nome baseado na versão
    exe_path = os.path.join(base_dir, "dist", f"{exe_name}.exe")
    exe_exists = False
    last_file_size = 0
    
    try:
        # Monitorar saída e atualizar progresso na MESMA linha
        for line in process.stdout:
            last_line = line.strip()
            
            # Atualizar estágio atual baseado na saída
            for pattern, progress in build_stages.items():
                if re.search(pattern, last_line):
                    current_progress = progress
                    current_stage = last_line.split("INFO: ")[1] if "INFO: " in last_line else last_line
            
            # Verificar se o arquivo EXE existe e está crescendo
            if current_progress >= 90 and os.path.exists(exe_path):
                exe_exists = True
                current_size = os.path.getsize(exe_path)
                if current_size > last_file_size:
                    last_file_size = current_size
                    current_stage = f"Criando executável: {current_size/1024/1024:.1f} MB"
            
            # Calcular tempos
            current_time = time.time()
            elapsed_time = current_time - start_time
            
            # Só atualiza a tela a cada 0.5 segundos para não sobrecarregar o terminal
            if current_time - last_update_time >= 0.5:
                last_update_time = current_time
                
                # Estimativa de tempo restante
                if current_progress > 0:
                    estimated_total_time = elapsed_time * 100 / current_progress
                    remaining_time = estimated_total_time - elapsed_time
                    
                    # Formatar tempos para exibição
                    elapsed_str = format_time(elapsed_time)
                    remaining_str = format_time(remaining_time)
                    
                    # Criar uma única linha de status
                    progress_bar = draw_progress_bar(current_progress)
                    status_line = f"{progress_bar} | {current_stage[:40]} | {elapsed_str}/{remaining_str}"
                    
                    # Limpar linha atual e escrever nova linha
                    sys.stdout.write('\r' + ' ' * 100)
                    sys.stdout.write('\r' + status_line)
                    sys.stdout.flush()
        
        # Esperar conclusão do processo
        process.wait()
        
    except KeyboardInterrupt:
        print("\n\nOperação cancelada pelo usuário.")
        try:
            process.terminate()
        except:
            pass
        return False
    
    # Conclusão e verificações finais
    elapsed_time = time.time() - start_time
    
    if process.returncode == 0:
        # Mostrar progresso completo na mesma linha
        sys.stdout.write('\r' + ' ' * 100)
        sys.stdout.write('\r' + f"{draw_progress_bar(100)} | Concluído! | {format_time(elapsed_time)}")
        sys.stdout.flush()
        print()  # Quebra de linha após conclusão
        
        if os.path.exists(exe_path):
            file_size = os.path.getsize(exe_path) / (1024*1024)
            print(f"\nExecutável criado com sucesso!")
            print(f"Tamanho: {file_size:.2f} MB")
            print(f"Executável disponível em: {exe_path}")
                
            # Instruções sobre como submeter o executável para revisão nos antivírus
            print("\n========== REDUZINDO DETECÇÕES DE FALSO POSITIVO ==========")
            print("Para evitar detecções incorretas como vírus, considere:")
            print("1. Enviar o executável para análise em: https://www.virustotal.com/")
            print("2. Se detectado como falso positivo, submeta para reavaliação em:")
            print("   - Microsoft: https://www.microsoft.com/en-us/wdsi/filesubmission")
            print("   - Avast/AVG: https://www.avast.com/false-positive-file-form.php")
            print("   - Kaspersky: https://www.kaspersky.com/submit-sample")
            print("   - Outros antivírus: visite o site do fabricante para submeter análise")
            print("============================================================")
            
            return True
        else:
            print("\nERRO: Executável não encontrado após compilação.")
            return False
    else:
        # Limpa a linha atual antes de mostrar erro
        sys.stdout.write('\r' + ' ' * 100)
        sys.stdout.write('\r')
        sys.stdout.flush()
        
        print(f"\nERRO: Compilação falhou com código {process.returncode}")
        
        # Sugerir execução como administrador se for um problema de permissão
        if "Acesso negado" in last_line or "PermissionError" in last_line:
            print("\nDetectado problema de permissão. Tente:")
            print("1. Fechar todas as instâncias do programa em execução")
            print("2. Executar este script como administrador")
            
            if not is_admin():
                print("\nDeseja tentar executar como administrador? (s/n): ", end='', flush=True)
                if input().lower() == 's':
                    script = os.path.abspath(sys.argv[0])
                    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, script, None, 1)
        
        if last_line:
            print(f"Última mensagem: {last_line}")
        return False

def format_time(seconds):
    """Formata segundos para exibição como minutos:segundos ou horas:minutos:segundos"""
    if seconds < 60:
        return f"{seconds:.1f}s"
    elif seconds < 3600:
        minutes = int(seconds // 60)
        seconds = seconds % 60
        return f"{minutes}m{int(seconds)}s"
    else:
        hours = int(seconds // 3600)
        seconds %= 3600
        minutes = int(seconds // 60)
        seconds = int(seconds % 60)
        return f"{hours}h{minutes}m{seconds}s"

if __name__ == "__main__":
    print("Sistema:", platform.system(), platform.release(), platform.architecture()[0])
    
    # Instala o PyInstaller se necessário
    try:
        import PyInstaller
        print(f"PyInstaller versão {PyInstaller.__version__} encontrado.")
    except ImportError:
        print("Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    try:
        # Constrói o executável
        build_exe_safe()
    except KeyboardInterrupt:
        print("\nOperação cancelada.")
    except Exception as e:
        print(f"Erro inesperado: {e}")
