import shutil
import time as t
def Backup():
    try:
        origem = r"C:\Users\MIKE\Desktop\automatizacaoExcell-main\automatizacaoExcell-main\dados.xlsx"
        destino = r"C:\Users\MIKE\Desktop"
        shutil.copy(origem, destino)
        print("Backup feito com sucesso!")
    except FileNotFoundError:
        print("\nErro ao procurar o arquivo\n")
    except PermissionError:
        print("\nSem permissão para acessar o arquivo ou pasta\n")
    except Exception as e:
        print(f"\nErro inesperado: {e}\n")
