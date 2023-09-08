import pandas as pd

from main import comecar
from verificarCaminhoNavWpp import verificarNav
import multiprocessing


wpps = pd.read_excel(r"C:\Users\Micro\Desktop\gerenciar_PLanilhas\informacoesWPP.xlsx")

processos = []


def main():
    if __name__ == '__main__':
        for i, wpp in enumerate(wpps['numero_padrao']):
            ondeMandarNumeros = str(wpps.loc[i, 'numero_padrao'])
            TextoMandar = str(wpps.loc[i, 'texto'])
            nomeWPP = str(wpps.loc[i, 'cobrador'])
            caminhoPlanilha = str(wpps.loc[i, 'caminhoPLan'])

            drive = "" #TODO: colocar o driver aqui  quando aprender a guardar ele

            p = multiprocessing.Process(target=comecar,
                                        args=(ondeMandarNumeros, TextoMandar, nomeWPP, caminhoPlanilha, drive))
            processos.append(p)

        for p in processos:
            p.start()

        for p in processos:
            p.join()


main()
#texte
