from tkinter import *
janela = Tk()
janela.title("Automatizar memorial")
janela.geometry("1000x1000")  # Defina o tamanho da janela (largura x altura)



#CLIENTE  
vcliente = Entry(janela, width=20)  # Defina a largura da caixa de entrada
vcliente.grid(row=0, column=1)  # Posicione a caixa de entrada na janela
label = Label(janela, text="Nome:")
label.grid(row=0, column=0)

#RG
vrg = Entry(janela, width=20)  # Defina a largura da caixa de entrada
vrg.grid(row=1, column=1)  # Posicione a caixa de entrada na janela
label = Label(janela, text="RG:")
label.grid(row=1, column=0) 





def obter_valor():
    nome = vcliente.get()
    rg = vrg.get()
    print(nome,rg)
   
    

botao = Button(janela, text="confirmar", command=obter_valor)
botao.grid(row=3, column=0, columnspan=2)  # Posicione o botão na janela
janela.mainloop()