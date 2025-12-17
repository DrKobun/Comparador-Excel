<h1> Baixador e Comparador de Planilhas Excel</h1>
<p> Este é o repositório de um programa para comparação de valores entre duas planilhas.</p>
<p> No momento atual, a comparação pode ser feita com as seguintes bases de dados:</p>
<ul>
  <li>SINAPI</li>
  <li>SICRO</li>
  <li>ORSE</li>
</ul>

<h2> Fluxo de uso do programa </h2>
<p> O fluxo de uso segue, mas não necessariamente, a seguinte ordem:</p>
<ul>
  <li> O usuário seleciona qual base de dados deseja fazer o download da planilha</li>
  <li> O usuário seleciona qual é o ano e o mês de download da planilha da base de dados</li>
  <li> O usuário seleciona qual tipo de planilha deseja baixar (Desonerado, Não Desonerado ou Ambos)</li>
  <li> O usuário seleciona qual tipo de arquivos deseja agrupar em um único arquivo Excel (Insumos, Sintéticos ou Ambos)</li>
  <li> O usuário seleciona quais estados deseja fazer o download. Podem ser baixados mais de um arquivo simultâneamente</li>
  <li> O usuário agrupa os arquivos baixados em um único arquivo Excel</li>
  <li> O usuário formata o arquivo Excel gerado após o agrupamento dos outros arquivos</li>
  <li> O usuário faz a comparação da planilha de projeto com a planilha agrupada</li>
</ul>

<p align="justify">Ao finalizar o download dos arquivos desejados, é hora de agrupá-los em um único arquivo Excel clicando no botão "Juntar". 
  Ao juntar os arquivos desejados, também é formatado o nome de cada arquivo, para que o usuário saiba o que cada planilha representa. A classificação de planilhas segue a seguinte regra:</p>
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
<h1>


<h2>Passo a passo de uso:</h2>
<h3>Para selecionar a base de dados:</h3>
<img src="/imagens-repo/selecao-base-dados.png" alt="Imagem demonstrando a seleção da base de dados do programa" height=500/>

<h3>Para selecionar a data e o mês desejados:</h3>
<img src="/imagens-repo/selecao-data.png" alt="Imagem demonstrando local de seleção de mês e ano do arquivo a ser baixado" height=500/>

<h3>Para selecionar o tipo de arquivo para baixar:</h3>
<img src="/imagens-repo/selecao-tipo-baixar.png" alt="Imagem demonstrando a seleção do tipo de arquivo que será baixado" height=500/>

<h3>Para selecionar os tipos de arquivo que serão agrupados:</h3>
<img src="/imagens-repo/selecao-tipo-juntar.png" alt="Imagem demonstrando os tipo de arquivos serão agrupados" height=500/>

<h3>Seleção de estados para serem baixados:</h3>
<img src="/imagens-repo/selecao-estados.png" alt="Imagem demonstrando local para seleção de estados para serem baixados" height=500/>

<h3>Botão para realizar o download dos arquivos selecionados:</h3>
<img src="/imagens-repo/selecao-baixar.png" alt="Imagem demonstrando local para download dos arquivos selecionados" height=500/>

<h3>Botão para agrupar arquivos em um único arquivo Excel:</h3>
<img src="/imagens-repo/selecao-juntar.png" alt="Imagem demonstrando local para agrupar arquivos das bases de dados baixados" height=500/>

<h3>Botão para formatar arquivos agrupados:</h3>
<img src="/imagens-repo/selecao-formatar.png" alt="Imagem demonstrando botão para formatar arquivos agrupados" height=500/>

<h3>Botão para selecionar arquivo Excel do projeto:</h3>
<img src="/imagens-repo/selecao-projeto.png" alt="Imagem demonstrando botão para selecionar arquivo Excel do Projeto" height=500/>

<h3>Botão para selecionar arquivo agrupado e formatado:</h3>
<img src="/imagens-repo/selecao-base-dados-comparar.png" alt="Imagem demonstrando botão para selecionar arquivo Excel agrupado e formatado" height=500/>


<h4>Observações:</h4>
<ul>
  <li>
    As planilhas da base de dados do ORSE estão armazenadas no Google Drive. Se forem ser baixadas usando a internet do Ministério, é necessário utilizar uma extensão de VPN no navegador. Recomendo a <b><a href="https://chromewebstore.google.com/detail/free-vpn-for-chrome-vpn-p/majdfhpaihoncoakbjgbdhglocklcgno?pli=1">VeePN</a></b> , usando específicamente a localidade: <b>USA - Virginia</b>.
  </li>
  <li>
    Para a comapração ser feita corretamente, é NECESSÁRIO que o ARQUIVO DO PROJETO, tenha uma planilha nomeada EXCLUSIVAMENTE como "Curva ABC", seguinto a ordem de colunas: Coluna A - Quantidade de itens OU Código do insumo ou seviço, Coluna B - Descrição, Coluna D - PREÇO UNITÁRIO.   
  </li>
</ul>
  
