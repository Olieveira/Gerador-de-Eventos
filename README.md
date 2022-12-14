<h1 align="center">Gerador de Eventos</h1>

<p align='center'>
  <img src="https://user-images.githubusercontent.com/107584427/207604012-331361e1-2d1a-4ce3-bc89-8fefc53868e8.png"/>
</p>

<p align="center">Esse gerador de eventos foi criado utilizando a linguagem de programação Visual Basic for Applications (VBA) e funcionalidades proporcionadas pelo Excel. Através de um formulário dinâmico, proporciona a possibilidade de gerar um cronograma enumerado de um evento, e a partir do mesmo gerar diversos tipos diferentes de arquivos, formatando as informações fornecidos pelo usuário de forma automática e simples!</p>

# Funcionalidades :exclamation:
 - `Geração dinâmica:` Através de um botão, é gerado objetos de formulário de forma dinâmica em tempo de execução, proporcionando uma versatilidade quanto à quantidade limite de linhas do cronograma!<br>
 - `Páginação:` Para possibilitar uma quantidade indefinida de itens e para controle do tamanho do formulário, é realizado uma limitação de linhas por página (8), que por sua vez, quando ultrapassado, ocasiona na geração de novas páginas, realizando a devida separação e exibição dos itens assim como controle de navegação entre páginas!<br>
 
<p align='center'>
  <img src='https://user-images.githubusercontent.com/107584427/207616830-ca60db58-6515-4083-b58f-6289cb1589c6.gif' height="350px" />
</p>
 
 - `Salvar como:` Através de funções associadas à botões, a planilha proporciona uma variedade de tipos de arquivos para salvar o cronograma.
 
 - - `HTML:` Para arquivos HTML, as colunas são formatadas com ajuste automático (extende a largura para ajustar o conteúdo), visto que a altura e quebra de linha das células da planilha não interfere no tamanho da linha da tabela em HTML:
 
 
 <p align='center'>
  <img src='https://user-images.githubusercontent.com/107584427/207620333-0166fc04-136e-4ac8-9989-e8b869ba5114.png' height="350px" />
</p>

 
 - - `PDF:` Em arquivos PDF, as colunas são formatadas com quebra de linhas para limitar a largura do documento, além de uma altura adicional em cada linha para melhor visibilidade das informações:
 

 <p align='center'>
  <img src='https://user-images.githubusercontent.com/107584427/207623054-e40501da-7da0-4588-a346-19f24a9b4604.png' height="350px" />
</p>


 - - `TXT:` Para arquivos TXT, a formatação realizada é a mesma de PDF, exceto que o cabeçalho é retirado, porém seu resultado varia conforme a quantidade de linhas e tamanho do texto dos itens:
 

 <p align='center'>
  <img src='https://user-images.githubusercontent.com/107584427/207654545-a2adb9ae-e1c9-4d26-a97e-1302b499614b.png' height="350px" />
</p>


 - - `DOC:` Na geração de arquivos do tipo DOC (word), devido a compatibilidade entre os programas do pacote Office, é realizado algumas configurações do documento, como alteração de margens, e depois é feito uma colagem especial das informações no formato de tabela, em seguida estilizando tanto o header do documento quanto as células da tabela, tudo via **VBA**: 


 <p align='center'>
  <img src='https://user-images.githubusercontent.com/107584427/207655994-46d11525-77b9-4085-a282-468569ccf987.png' height="400px" />
</p>


 - - `XLSX:` Na geração em XLSX (planilha excel), é realizado a configuração de impressão e a estilização das linhas e colunas para que o conteúdo fique devidamente encaixado na àrea de impressão: 


 <p align='center'>
  <img src='https://user-images.githubusercontent.com/107584427/207657922-b5319659-0f6f-4561-a255-1f5fc0d97b4a.png' height="400px" />
</p>

 - `Imprimir:` Uma planilha é criada e adaptada no modelo XLSX. Com as informações estilizadas e a página configurada é realizada a solicitação da impressão.

 - `Importar:` Ao clicar no botão importar, o programa solicita que o usuário selecione um arquivo XLSX, o qual verifica a compatibilidade para importação baseado no modelo XLSX. Com um arquivo válido selecionado, o programa irá importar todas as informações da planilha selecionada.


<p align='center'>
  <img src='https://user-images.githubusercontent.com/107584427/207663704-064fef4a-98e2-45f2-a055-1d92629d6b53.gif' height="350px" />
</p>


# Características do projeto :hammer:

- `FORMULÁRIOS:` A planilha utiliza um formulário como principal meio de comunicação entre o usuário e o programa, usufruindo das funcionalidades oferecidas pelo excel e automatizando a geração de objetos em tempo de execução para proporcionar uma interface intuitiva e agradável.<br>

- `BASE DE DADOS:` É utilizado uma aba oculta como base de dados simples para controle de informações, como quantidade de objetos criados, quantidade de páginas e página atual, assim como os valores padrão dos ListBox's.

- `CONTROLE DE ARQUIVOS:` Ao gerar um arquivo, antes do mesmo ser aberto e/ou salvo, é verificado se a pasta destino padrão existe, caso não exista, a mesma é criada, assim como uma subpasta para cada extensão de arquivo que posteriormente será utilizada para organização dos arquivos gerados. Essa pasta é criada no mesmo diretório onde a planilha foi executada!

- `AVISOS:` Em várias das funções criadas, foram adicionados tratativas de erros, que avisam o usuário de um erro e/ou orientam à solução, evitando mensagens de erro de sistema nativas do excel.


# Tecnologias Utilizadas :white_check_mark:

 - `Microsoft Visual Basic for Applications 7.1`
 - `Microsoft 365`
 - `Microsoft Excel`
 - `Microsoft Word`

# Objetivo :dart:
Programa criado com intuito de padronizar arquivos em processos acadêmicos/profissionais/pessoais, assim como expandir conhecimento nas àreas das tecnologias utilizadas, com ênfase em conceitos como programação orientada a objetos, classes, funções, parâmetros obrigatórios, parâmetros opcionais, tipagem de variáveis, manipulação de arquivos, tratativas de erros, criação de objetos em tempo de execução, lógica computacional, boas práticas de programação etc. :student:


# Autor

[<img src="https://user-images.githubusercontent.com/107584427/205271408-568fcb74-3afe-42b1-a16b-c53a3922bf86.jpg" width=115><br><sub>Nathan de Oliveira de Melo</sub>](https://github.com/Olieveira)
