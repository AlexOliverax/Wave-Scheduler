---
description: Lançar uma nova versão (Release) do Waves Scheduler
---
Este workflow automatiza o processo de atualizar a versão do aplicativo, comitar o código e disparar o auto-update usando GitHub Actions.

Sempre que o usuário pedir para "lançar uma nova versão", "atualizar o app", ou "criar uma release", siga exatamente estes passos. A tag `// turbo-all` no final indicará para você executar os comandos de terminal permitidos automaticamente.

### Passos

1. Pergunte ao usuário qual deve ser o número da nova versão (ex: `2.1.0`). Só prossiga com o número correto e confirmado.
2. Atualize o arquivo `VERSION` na raiz do projeto com o número exato da nova versão.
3. Atualize a linha `#define AppVersion "X.X.X"` dentro de `WavesScheduler_Setup.iss` para refletir a nova versão.
4. Execute os seguintes comandos para comitar, criar a tag e realizar o push. Substitua `<NOVA_VERSAO>` pela versão escolhida.

// turbo-all
```powershell
git add VERSION WavesScheduler_Setup.iss
git commit -m "chore: release v<NOVA_VERSAO>"
git tag v<NOVA_VERSAO>
git push origin main
git push origin v<NOVA_VERSAO>
```

5. Informe ao usuário que o comando foi disparado. A GitHub Action configurada no repositório assumirá o processo a partir daqui, compilando o PyInstaller, gerando o Inno Setup Installer, e publicando a nova Release no GitHub. Os usuários finais vão baixar essa atualização sozinhos da próxima vez que abrirem o aplicativo.
