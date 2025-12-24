<h1> Baixador e Comparador de Planilhas Excel</h1>
<p> Este é o repositório de um programa para comparação de valores entre duas planilhas.</p>
<p> No momento atual, a comparação pode ser feita com as seguintes bases de dados:</p>
<ul>
  <li><a href="https://www.caixa.gov.br/site/paginas/downloads.aspx">SINAPI</a> (Sistema Nacional de Pesquisa de Custos e Índices da Construção Civil)</li>
  <li><a href="https://www.gov.br/dnit/pt-br/assuntos/planejamento-e-pesquisa/custos-referenciais/sistemas-de-custos/sicro/relatorios/relatorios-sicro">SICRO</a> (Sistema de Custos Referenciais de Obras)</li>
  <li><a href="https://orse.cehop.se.gov.br/">ORSE</a> (Sistema de Orçamento de Obras de Sergipe)</li>
</ul>

<br>

[![Download](https://img.shields.io/badge/GitHub-Download-green?logo=github&logoColor=white)](https://github.com/DrKobun/Comparador-Excel/raw/refs/heads/main/sinapi-gui-app/src/dist/main.exe)


</p>

<h2> Fluxo de uso do programa </h2>
<p> O fluxo de uso segue, mas não necessariamente, a seguinte ordem:</p>
<ul>
  <li> O usuário seleciona qual base de dados deseja fazer o download da planilha.</li>
  <li> O usuário seleciona qual é o ano e o mês de download da planilha da base de dados.</li>
  <li> O usuário seleciona qual tipo de planilha deseja baixar (Desonerado, Não Desonerado ou Ambos).</li>
  <li> O usuário seleciona qual tipo de arquivos deseja agrupar em um único arquivo Excel (Insumos, Sintéticos ou Ambos).</li>
  <li> O usuário seleciona quais estados deseja fazer o download. Podem ser baixados mais de um arquivo simultâneamente.</li>
  <li> O usuário agrupa os arquivos baixados em um único arquivo Excel.</li>
  <li> O usuário formata o arquivo Excel gerado após o agrupamento dos outros arquivos.</li>
  <li> O usuário faz a comparação da planilha de projeto com a planilha agrupada.</li>
</ul>

<p align="justify">Ao finalizar o download dos arquivos desejados, é necessário agrupá-los em um único arquivo Excel clicando no botão "Juntar". 
  Após juntar os arquivos desejados, é necessário formatar a planilha agrupada, clicando no botão <b>Formatar</b>. 
  <br>
  Com o arquivo gerado, é possível fazer a comparação com a planilha do projeto. 
  <br>
  <br>
  A classificação de planilhas segue a seguinte regra:</p>
<h1>
  <h3> Prefixos das bases de dados: </h3>
  <ul>
    <li>
      SINA = SINAPI   
    </li>
    <li>
      SICRO = SICRO
    </li>
    <li>
      ORS = ORSE
    </li>
  </ul>
<h1>
<h3> Tipos de dados: </h3>
<ul>
  <li>
    SIN = Sintéticos
  </li>
  <li>
    SER = Serviços
  </li>
  <li>
    INS = Insumos
  </li>
</ul>
<h1>
  <h3> Para planilhas SICRO: </h3>
  <ul>
    <li>
      COMP = Composições
    </li>
    <li>
      EQP = Equipamentos
    </li>
    <li>
      DES = Desonerado (EQP)
    </li>
    <li>
      MAT = Materiais
    </li>
  </ul>
<h1>
  <h3> Tipos de valores das planilhas: </h3>
  <ul>
    <li>
      DES = Desonerado
    </li>
    <li>
      NDS = Não Desonerado
    </li>
    <li>
      AMB = Ambos
    </li>
  </ul>
  
<p>Exemplo de planilha nomeada: <b>SINA-SIN-PB-202006-DES</b></p>
<p>O exemplo significa que a planilha se refere ao SINAPI, Sintéticos, Paraíba, ano de 2020, mês de Junho, e é Desonerada.</p>


<h2>Passo a passo de uso:</h2>
<h3>Para selecionar a base de dados:</h3>
<img src="/imagens-repo/selecao-base-dados.png" alt="Imagem demonstrando a seleção da base de dados do programa" height=500/>

<h3>Para selecionar o ano e o mês desejados:</h3>
<img src="/imagens-repo/selecao-data.png" alt="Imagem demonstrando local de seleção de mês e ano do arquivo a ser baixado" height=500/>

<h3>Para selecionar o tipo de arquivo para baixar:</h3>
<img src="/imagens-repo/selecao-tipo-baixar.png" alt="Imagem demonstrando a seleção do tipo de arquivo que será baixado" height=500/>

<h3>Para selecionar os tipos de arquivo que serão agrupados:</h3>
<img src="/imagens-repo/selecao-tipo-juntar.png" alt="Imagem demonstrando os tipo de arquivos serão agrupados" height=500/>

<h3>Seleção de estados para serem baixados:</h3>
<img src="/imagens-repo/selecao-estados.png" alt="Imagem demonstrando local para seleção de estados para serem baixados" height=500/>

<h3>Botão para realizar o download dos arquivos selecionados:</h3>
<img src="/imagens-repo/selecao-baixar.png" alt="Imagem demonstrando local para download dos arquivos selecionados" height=500/>

<h3>Botão para agrupar planilhas baixadas em um único arquivo Excel:</h3>
<img src="/imagens-repo/selecao-juntar.png" alt="Imagem demonstrando local para agrupar arquivos das bases de dados baixados" height=500/>

<h3>Botão para formatar arquivos agrupados:</h3>
<img src="/imagens-repo/selecao-formatar.png" alt="Imagem demonstrando botão para formatar arquivos agrupados" height=500/>

<h3>Botão para selecionar arquivo Excel do projeto:</h3>
<img src="/imagens-repo/selecao-projeto.png" alt="Imagem demonstrando botão para selecionar arquivo Excel do Projeto" height=500/>

<h3>Botão para selecionar arquivo agrupado e formatado:</h3>
<img src="/imagens-repo/selecao-base-dados-comparar.png" alt="Imagem demonstrando botão para selecionar arquivo Excel agrupado e formatado" height=500/>

<h3>Botão para iniciar a comparação:</h3>
<img src="/imagens-repo/selecao-iniciar-comparacao.png" alt="Imagem demonstrando botão para iniciar a comparação entre dois arquivos selecionados" height=500/>

<h3>Para apagar dados das pastas do fluxo de funcionamento do programa:</h3>
<img src="/imagens-repo/selecao-apagar.png" alt="Imagem demonstrando botão para apagar pastas do fluxo de funcionamento do programa, para ser feito uma nova comparação" height=500/>

<p>Após serem escolhidos os arquivos para ser feita a comparação, o algoritmo procura na pasta de projetos, se existe alguma planilha nomeada como "Curva ABC", se o encontrar, é feita a comparação de todas as linhas da coluna B (Descrição) e todas as linhas da coluna D (Valor). O resultado da comparação é gerado em uma nova planilha nomeada como "Resultado da Comparação".</p>
<br>
<p>As cores das linhas representam:</p>

<ul>
  <li><img src="https://img.shields.io/badge/Verde-green" /> -> O valor do item da Curva ABC está <b>MENOR</b> do que a planilha da base de dados.</li>
  <li><img src="https://img.shields.io/badge/Vermelho-red" /> -> O valor do item da Curva ABC está <b>MAIOR</b> do que a planilha da base de dados.</li>
  <li><img src="https://img.shields.io/badge/Azul-blue" /> -> O valor do item da Curva ABC está <b>IGUAL</b> à planilha da base de dados.</li>
  <li><img src="https://img.shields.io/badge/Cinza-gray" /> -> O item da Curva ABC <b>não foi encontrado</b> em nenhuma planilha das bases de dados.</li>
</ul>



<h4>Observações:</h4>
<ul>
<li align="justify">
    As planilhas da base de dados do <b>ORSE</b> estão armazenadas no <b>Google Drive</b>. Se forem baixadas usando a internet do Ministério, é necessário utilizar uma extensão de VPN no navegador. Recomendo a <a href="https://chromewebstore.google.com/detail/free-vpn-for-chrome-vpn-p/majdfhpaihoncoakbjgbdhglocklcgno?pli=1" target="_blank">VeePN ↗</a>, usando específicamente a localidade: <b>USA - Virginia</b>. 
<br>
<br>
<b>Segue um rápido passo a passo para instalar a extensão no navegador (Google Chrome):</b>
<h3>Para instalar:</h3>
<img src="/imagens-repo/VeePN.png"/>
<h3>Para adicionar:</h3>
<img src="/imagens-repo/VeePN-adicionar.png"/>
<h3>Para fixar:</h3>
<img src="/imagens-repo/VeePN-fixar.png"/>
<h3>Para selecionar a região (USA - Virginia):</h3>
<img src="/imagens-repo/VeePN-localidade.png" height="500"/>
<img src="/imagens-repo/VeePN-virginia.png" height="500"/>
<h3>Para iniciar conexão com VPN:</h3>
<img src="/imagens-repo/VeePN-iniciar.png" height="500"/>
</li>
<li align="justify">
    Para a comparação ser feita corretamente é NECESSÁRIO que o ARQUIVO DO PROJETO tenha uma planilha nomeada EXCLUSIVAMENTE como "Curva ABC" (O nome da planilha precisa respeitar as maiúsculas e minúsculas desse exemplo), seguindo a ordem de colunas: 
<br>
<br>
<ul>
  <li><b>Coluna A</b> - Quantidade de itens OU Código do insumo ou seviço</li>
  <li><b>Coluna B</b> - Descrição</li>      
  <li><b>Coluna D</b> - Preço Unitário</li>
</ul>
<p>Exemplo de ordem de colunas da planilha <b>"Curva ABC"</b> (da planilha do <b>PROJETO</b>):</p>      
<img src="/imagens-repo/excel-exemplo.PNG" />
  
<li>Antes de se fazer um download da base de dados do <b>SINAPI</b>, é necessário abrir uma janela do navegador do site de download, caso contrário o download é negado.</li>
<br>
<b><p>Para abrir o link de downloads do SINAPI rapidamente:</p></b>
<img src="/imagens-repo/SINAPI-site.png" alt="Imagem demonstrando local para abrir rapidamente o site de downloads do SINAPI" height=500/>
<br>
<b><p align="justify">Caso os downloads ainda não funcionem após abrir o site de downloads do SINAPI, reabra o programa (mantendo o site do SINAPI aberto), e tente baixar os arquivos desejados novamente.</p></b>
  <li>Para fazer uma nova comparação, ou gerar um novo arquivo agrupado e formatado, clique no botão "apagar dados", serão apagados quaisquer arquivos que estejam nas pastas de funcionamento do programa, apagar esses arquivos para fazer uma nova comparação é necessário, pois o programa leva em consideração quaisquer arquivos Excel que estejam na pasta do fluxo de funcionamento.</li>
</ul>
  
