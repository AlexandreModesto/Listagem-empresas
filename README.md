# Listagem-empresas
Esta é uma automação com ```Pandas``` e ```PySimpleGUI```, para que formatar uma planilha excel para um .txt com certas pré definições, como caracteres especiais e quantidade de linhas máximas. A função do ```PySimpleGUI``` é para a facilitação do carregamento do arquivo a ser formatado, fazendo com que não seja necessário ele estar no mesmo diretório onde se encontra o código-fonte ou o arquivo .exe

o método ```open()``` é para abrir um arquivo ou criar um com o nome passado no parametro e o parametro ```'w'``` é para escrever no arquivo.
a váriavel ```indice``` é necessária para mencionar qual o index que esta sendo interado no método ```df.iterrows()```
![image](https://user-images.githubusercontent.com/59314251/223131064-63d77d5e-6260-4b49-85f0-45517a7a9872.png)

```sg.popup_get_file()``` é um método onde ele abre um pop-up para upload de arquivo 

![image](https://user-images.githubusercontent.com/59314251/223133822-de53a651-e230-4021-b8b6-8b272769b18c.png)

a primeira linha abaixo faz com que sempre que eu mencione a váriavel ```concatenando```, o código saiba que é na coluna passada dentro do ```df[]``` e no index passado dentro do ```values[]```. A conversão dessa variavel para string é para poder posteriormente acessar cada caractere dentro da célula

![image](https://user-images.githubusercontent.com/59314251/223134433-7174c606-a4a8-4b92-9ae2-0977ca63e154.png)

Abaixo é uma árvore de decisão onde checa se é CPF ou CNPJ. O sistema que irá lê esse arquivo só entende CPF/CPNJ com alguns zeros antes do numero ou o valor 1. Há uma checagem de decisão para caso os CNPJ/CPF comecem com zero

![image](https://user-images.githubusercontent.com/59314251/223136886-d29e869e-bdb9-4939-a40d-81724743c33a.png)

Abaixo é algo recorrente no código, é para caso na planilha fonte haja caracteres especiais da lingua portuguesa eles possam ser convertidos para a formatação adequada

![image](https://user-images.githubusercontent.com/59314251/223138169-5eb85d2e-9d73-4be4-9e99-a3e4fde24416.png)

Por fim, ele pega todas as váriavies já formatadas com os valores finais concatenadas em uma variavel. O método ```write()``` para escrever no arquivo .txt o que foi passado em seu parâmetro

![image](https://user-images.githubusercontent.com/59314251/223139158-dfe5b533-e267-464d-9b16-fe88ddb78eef.png)
