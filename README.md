<h1 align="left">📄 Conversor Universal para PDF</h1>

<p align="left">
  Ferramenta em Python com interface gráfica para transformar diversos arquivos em PDF.<br>
  Suporta documentos do Office (Word, Excel, PowerPoint) e imagens (JPG, PNG).
</p>

<hr>

<h2>🔍 O que é</h2>

<p>
<code>universal-pdf-converter</code> é um programa prático para quem precisa unificar documentos ou converter ficheiros do Office sem abrir programa por programa.<br>
Com ele, você pode:
</p>

<ol>
  <li>Selecionar múltiplos documentos Word, Excel e PowerPoint de uma vez</li>
  <li>Selecionar várias fotos para criar um único álbum em PDF</li>
  <li>Converter tudo automaticamente com apenas um clique</li>
  <li>Salvar os resultados diretamente na pasta de Downloads</li>
</ol>

<hr>

<h2>📂 Estrutura</h2>

<pre><code>├── conversor_pdf.py    <-- programa principal (Interface GUI)
├── requirements.txt    <-- lista de dependências necessárias</code></pre>

<hr>

<h2>⚙️ Funcionalidades</h2>

<p>
O script utiliza bibliotecas poderosas para garantir a qualidade:
</p>
<ul>
  <li><b>Office:</b> Usa o motor do Word/Excel/PPT instalado para garantir que nada saia do lugar.</li>
  <li><b>Imagens:</b> Usa a biblioteca <code>Pillow</code> para unir fotos em alta qualidade.</li>
  <li><b>Interface:</b> Construída em <code>Tkinter</code> para ser leve e funcional.</li>
  <li><b>Automação:</b> Deteta automaticamente o caminho de Downloads do seu computador.</li>
</ul>

<hr>

<h2>🛠️ Instalação</h2>

<ol>
  <li>Instale o Python 3</li>
  <li>Certifique-se de que tem o Microsoft Office instalado (necessário para arquivos .docx, .xlsx e .pptx)</li>
  <li>Instale as dependências necessárias executando no terminal:
    <pre><code>pip install Pillow docx2pdf comtypes</code></pre>
  </li>
</ol>

<hr>

<h2>🚀 Como usar</h2>

<ol>
  <li>Abrir o terminal ou a sua lista de ferramentas</li>
  <li>Executar o ficheiro:
    <pre><code>python conversor_pdf.py</code></pre>
  </li>
  <li>No ecrã que abrir, clique em <b>"Adicionar Arquivos"</b></li>
  <li>Se for unir imagens, dê um nome ao ficheiro no campo indicado</li>
  <li>Clique em <b>"CONVERTER PARA PDF"</b> e aguarde o aviso de sucesso</li>
</ol>

<hr>

<h2>💡 Possíveis melhorias</h2>

<ul>
  <li>Adicionar suporte para ficheiros de texto simples (.txt)</li>
  <li>Implementar a função de arrastar e soltar (Drag and Drop)</li>
  <li>Opção para comprimir o PDF final para ocupar menos espaço</li>
  <li>Conversão de PDFs de volta para outros formatos</li>
</ul>
