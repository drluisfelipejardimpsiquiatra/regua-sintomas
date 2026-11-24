# Rumo Livre · fotos dos testes: como instalar

Faça isto uma vez, no computador. Leva uns 5 minutos.
Até instalar, a foto do teste fica guardada só no celular do paciente. O painel mostra "foto no celular" e o senhor pode pedir para ver na consulta.

## Passo a passo
1. Abra https://script.google.com e clique em **Novo projeto**.
2. Dê o nome **Rumo Livre · fotos**. Apague o que estiver no editor, cole todo o conteúdo de `rumo-fotos.gs` e clique em **Salvar**.
3. Engrenagem **Configurações do projeto** → **Fuso horário**: *(GMT-03:00) Horário de Brasília*.
4. No editor, escolha a função **instalar** e clique em **Executar**. Autorize com a sua conta (Avançado → Acessar → Permitir).
   O registro deve mostrar: **"Instalado. Pasta: Rumo Livre · fotos dos testes"**.
5. **Implantar → Nova implantação** → tipo **App da Web**:
   - *Executar como*: **Eu**;
   - *Quem pode acessar*: **Qualquer pessoa**.
   Clique em **Implantar** e copie o **URL do app da Web** (termina em `/exec`).
6. Mande esse endereço para o Claude. Ele coloca o endereço em `FOTO_URL`, no `rumo.html`, e publica.

## Segurança e LGPD
- As fotos ficam no **seu** Google Drive, numa pasta privada. Ninguém mais tem acesso.
- O nome do arquivo leva só o código pessoal do link (`RL_ABCD2345_2026-11-26_073012.jpg`), nunca nome ou CPF.
- O endereço `/exec` aceita apenas imagens JPEG com código e data válidos. Ele não lê nada do Drive.
- Para trocar o endereço (por exemplo, se vazar), faça uma nova implantação e arquive a antiga.
