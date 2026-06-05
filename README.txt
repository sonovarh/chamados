LEITOR DE CHAMADOS RH SONOVA - V6

Como usar:
1. Extraia todo o ZIP em uma pasta.
2. Mantenha o arquivo Excel na mesma pasta do leitor:
   Chamados RH Sonova (respostas) até 28-05-2026.xlsx
3. Clique duas vezes em abrir_leitor.bat.
4. O navegador será aberto em localhost e o dashboard buscará os dados reais do Excel pela rota /api/workbook.

Correção desta versão:
- A leitura automática não depende mais da biblioteca XLSX online.
- O arquivo abrir_leitor.bat inicia um server.py local.
- O server.py lê o .xlsx real usando apenas bibliotecas padrão do Python.
- Se você substituir a planilha por uma versão atualizada com o mesmo nome, o leitor passa a ler a nova versão.
- Se abrir index.html direto, o navegador pode usar os dados embutidos como fallback, mas a leitura real automática exige o .bat.

Funcionalidades mantidas:
- Máximo de três cores principais.
- Filtro clicando nos cards.
- Filtro clicando nas classificações.
- Glossário geral de chamados.
- Classificadores com grupo, prioridade, confiança e termos encontrados.
- SLA calculado como abertura + 2 dias.
- Exportação CSV filtrada.

Anderson Marinho | Igarapé Digital
