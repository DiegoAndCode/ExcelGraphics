import pandas as pd
import matplotlib.pyplot as plt
import datetime

class Graficos:
    def __init__(self):
        self.df = None
        self.arquivo = 'clientes_exemplo.xlsx'
        self.titulo_eixo = 'Regime do Cliente'
        self.titulo_grafico = 'Gráfico de Regimes de Clientes'

    def ler_arquivo(self):
        try:
            self.df = pd.read_excel(self.arquivo)
        except FileNotFoundError:
            print(f"O arquivo '{self.arquivo}' não foi encontrado.")
            exit()

    def data_geracao(self):
        data_geracao = datetime.date.today().strftime("%d/%m/%Y")
        plt.figtext(0.98, 0.01, f"Data de geração: {data_geracao}", ha="right", fontsize=8)

    def grafico_barras(self, index, values):
        plt.bar(index, values)
        plt.xlabel(self.titulo_eixo)
        plt.ylabel('Quantidade')
        plt.title(self.titulo_grafico)

        for i, v in enumerate(values):
            plt.text(i, v + 0.1, str(v), ha='center')

        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()

        self.data_geracao()

        plt.subplots_adjust(bottom=0.4)

        nome_arquivo = f'grafico_barras.png'
        plt.savefig(nome_arquivo)
        plt.close()

    def grafico_barras_horizontal(self, index, values):
        plt.barh(index, values)
        plt.xlabel('Quantidade')
        plt.ylabel(self.titulo_eixo)
        plt.title(self.titulo_grafico)

        for i, v in enumerate(values):
            plt.text(v + 0.1, i, str(v), va='center')

        plt.tight_layout()

        self.data_geracao()

        plt.subplots_adjust(bottom=0.15)

        nome_arquivo = f'grafico_barrash.png'
        plt.savefig(nome_arquivo)
        plt.close()

    def grafico_linhas(self, index, values):
        plt.plot(index, values, marker='o', linestyle='-')

        plt.xlabel(self.titulo_eixo)
        plt.ylabel('Quantidade')
        plt.title(self.titulo_grafico)

        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()

        self.data_geracao()

        plt.subplots_adjust(bottom=0.4)

        nome_arquivo = f'grafico_linhas.png'
        plt.savefig(nome_arquivo)
        plt.close()

    def grafico_pizza(self, index, values):
        plt.pie(values, labels=index, autopct='%1.1f%%', startangle=140)
        plt.axis('equal')

        plt.title(f'{self.titulo_grafico}\n')
        plt.tight_layout()

        self.data_geracao()

        plt.subplots_adjust(bottom=0.08)

        nome_arquivo = f'grafico_pizza.png'
        plt.savefig(nome_arquivo)
        plt.close()

    def gerar_graficos(self):
        self.ler_arquivo()
        contagem = self.df['REGIME'].value_counts()
        contagem_acendente = contagem.sort_values(ascending=True)

        self.grafico_linhas(contagem.index, contagem.values)
        self.grafico_barras(contagem.index, contagem.values)
        self.grafico_barras_horizontal(contagem_acendente.index, contagem_acendente.values)
        self.grafico_pizza(contagem.index, contagem.values)

if __name__ == '__main__':
    graficos = Graficos()
    graficos.gerar_graficos()
