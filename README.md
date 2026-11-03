# CNG · Controle de Postagens

Sistema para a equipe do CNG registrar o que foi postado (Instagram, TikTok, YouTube e Site) e acompanhar a grade e as metas do dia. Roda como app web do Google Apps Script, com uma planilha do Google como banco de dados. A equipe entra sem login, só com nome + PIN.

## Regras padrão

| Rede | Regra | Dia útil | Fim de semana / feriado |
|---|---|---|---|
| Instagram | Grade por horário | 1 em 1 hora (07h–22h) | 2 em 2 horas (08h–22h) |
| TikTok | Meta diária | 8 posts | 8 posts |
| YouTube | Meta diária | 4 posts | 4 posts |
| Site | Meta diária | 5 posts | 3 posts |

Equipe: Thais, Iones, Yan, Vitor (meta de 3 vídeos/dia) e Bruno (trabalhos extras). Tudo isso pode ser alterado na tela **Ajustes**, entrando como Gestor.

## Arquivos

- `Code.gs` — servidor (lê e grava na planilha, confere PINs e a senha do gestor)
- `Index.html` — a tela do sistema

## Instalação

1. Crie uma planilha do Google e abra **Extensões → Apps Script**.
2. Cole o conteúdo de `Code.gs` no arquivo `Código.gs`.
3. Na linha 7, troque `troque-esta-senha` pela senha do gestor. **Não envie a senha real para este repositório.**
4. Crie um arquivo HTML chamado `Index` e cole o conteúdo de `Index.html`.
5. Salve, escolha a função `setup` e clique em **Executar**. Depois, autorize o acesso.
6. **Implantar → Nova implantação → App da Web**. Em "Executar como", escolha **Eu**. Em "Quem pode acessar", escolha **Qualquer pessoa**.
7. Envie à equipe o link que termina em `/exec`.

Para atualizar o sistema, cole os arquivos novos, salve e vá em **Implantar → Gerenciar implantações → editar → Nova versão**. O link continua o mesmo.

## Planilha

- **Registros**: uma linha por registro (data, horário, atividade, redes, tipo, pauta, responsável, link, status…)
- **Config**: ajustes do sistema (não editar à mão)
- **PINs** (oculta): PINs codificados da equipe
