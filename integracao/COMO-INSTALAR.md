# Pré-consulta → prontuário do V Saúde: como instalar

Faça isto uma vez, no computador (no celular o editor do Google não funciona bem). Leva uns 10 minutos.

## Antes de começar
- **Chave da API do V Saúde.** É a mesma chave que já está sendo usada. Se não tiver em mãos, peça ao suporte do V Saúde a "API Key da API pública".
- Abra a planilha **"Pré-consulta por voz · recebimento (respostas)"** no Google Drive.

## Passo a passo
1. Na planilha, menu **Extensões → Apps Script**.
2. Apague o que estiver no editor, cole todo o conteúdo do arquivo `vsaude-preconsulta.gs` e clique em **Salvar** (ícone de disquete).
3. À esquerda, clique na engrenagem **Configurações do projeto**:
   - em **Fuso horário**, escolha *(GMT-03:00) Horário de Brasília*;
   - no fim da página, em **Propriedades do script**, clique em **Adicionar propriedade**:
     nome `VSAUDE_API_KEY`, valor = a chave do V Saúde. Salve.
4. Volte ao editor (ícone `< >`). No seletor de funções, no alto, escolha **instalar** e clique em **Executar**.
5. O Google vai pedir autorização. Escolha a sua conta → **Avançado** → **Acessar (não seguro)** → **Permitir**.
   O aviso aparece porque o programa é seu e não passou por revisão do Google; ele só acessa esta planilha, o V Saúde e o envio de e-mail de aviso para você.
6. No registro de execução deve aparecer: **"Instalado. A chave do V Saúde funciona."**

## Como funciona depois
- Cada pré-consulta enviada é anexada como PDF na pasta de **Arquivos** do paciente no V Saúde, com o nome
  `Pré-consulta – DD-MM-AAAA.pdf`, em poucos segundos.
- A planilha ganha a coluna **"Status V Saúde"**:
  - ✅ anexado (com o nº do paciente);
  - ⚠ CPF não encontrado, ou mais de um cadastro com o mesmo CPF. Você recebe um e-mail e o programa tenta de novo a cada hora. Basta corrigir o cadastro no V Saúde que ele anexa sozinho.
- Para não tentar mais uma linha, escreva ⛔ no começo do status.

## Teste recomendado
1. A secretária cadastra no V Saúde um paciente **TESTE** com o CPF **529.982.247-25**.
2. Responda a pré-consulta pelo link usando esse CPF.
3. Confira a coluna "Status V Saúde" e a pasta de Arquivos do paciente TESTE.
4. Apague o paciente TESTE e a linha da planilha.
