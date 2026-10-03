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

Novidades V7 - Maiores acionadores:
- Painel "Maiores acionadores": ranking de quem mais abre chamados, agrupado por CPF
  (sem CPF: herda o CPF pelo mesmo e-mail; senão agrupa por e-mail/nome).
- CPF normalizado (recupera zero à esquerda perdido no Excel) e validação do dígito verificador.
- Filtros do painel: Top N, mínimo de chamados, só com chamado em aberto / fora do SLA / sem CPF, busca.
- Seleção por caixa (linha amarela); sem seleção, exporta a lista exibida.
- Exportar contatos (Excel): abas Contatos (CPF, Nome, E-mail), Acionadores_Detalhe e Chamados.
- Exportar CSV (UTF-8 com BOM, separador ;) e Copiar e-mails (separados por ; para Outlook/Teams).
- Clique no nome filtra o dashboard pela pessoa; duplo clique abre o histórico completo com exportação individual.
- server.py: corrigida a leitura de planilhas cujo caminho interno das abas é absoluto (/xl/worksheets/...).

Anderson Marinho | Igarapé Digital
