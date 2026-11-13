# Busca Aluguel

Programa que procura apartamentos para alugar no Kijiji, Centris, Appartogo (inclui LogisQuébec) e Facebook Marketplace, aplica os seus critérios e gera um Excel só com os anúncios novos.

## Como abrir no Mac (primeira vez)

1. Baixe o arquivo do seu Mac na página de downloads:
   - Mac com chip Apple (M1, M2, M3, M4): `BuscaAluguel-Mac-Apple.zip`
   - Mac com Intel: `BuscaAluguel-Mac-Intel.zip`
   - Para saber qual é o seu: menu Apple > Sobre este Mac.
2. Dê dois cliques no .zip e arraste o "Busca Aluguel" para a pasta Aplicativos.
3. Abra o app. O Mac vai avisar que não consegue verificar o desenvolvedor. Clique em OK.
4. Vá em Ajustes do Sistema > Privacidade e Segurança, role até o fim e clique em **Abrir Mesmo Assim**.
5. Pronto. Das próximas vezes o app abre normalmente.

O app precisa do Google Chrome instalado (google.com/chrome).

## Como usar

1. Ajuste os critérios na janela (cidades, preço, tamanho, área, idade do prédio, garagem).
2. Clique em **Buscar**. Uma janela do Chrome abre e percorre os sites sozinha. Não feche essa janela.
3. Na primeira vez, se o Facebook pedir login, faça o login nessa janela do Chrome.
4. No final, o app oferece abrir o Excel.

Os resultados ficam em Documentos > Busca Aluguel. O Excel tem 3 abas:
- **Atende**: todos os critérios confirmados no anúncio.
- **Verificar**: nada reprovado, mas falta alguma informação no anúncio.
- **Excluídos**: reprovou em algum critério (com o motivo).

Cada busca mostra só anúncios que ainda não apareceram. Para ver tudo de novo, clique em **Limpar histórico**.

## Limitações

- As buscas funcionam para cidades da Grande Montreal. No Appartogo, só para as cidades da lista conhecida (Lanaudière, Laval, Montreal, Rive-Sud e Laurentides principais).
- Os sites mudam com o tempo; se algum parar de trazer resultados, o programa precisa de ajuste.
