import shutil
import time as t
try:
    origem = r"\\DESKTOP-VJ8B4AG\impressão digital\EXECELL\dados.xlsx"
    destino = r"C:\Users\JESUS TE AMA\Desktop\backup"
    shutil.copy(origem, destino)
    print("Backup feito com sucesso!")
except FileNotFoundError:
    print("\nErro ao procurar o arquivo\n")
except PermissionError:
    print("\nSem permissão para acessar o arquivo ou pasta\n")
except Exception as e:
    print(f"\nErro inesperado: {e}\n")
