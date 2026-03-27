# ambiente-check

Script PowerShell para coleta e validação de ambiente para implantação do **MobyPharma / MobyCRM (FCerta/Phusion)**.

---

## Como usar

Cole o comando abaixo no **PowerShell** da máquina do cliente:

```powershell
irm "https://raw.githubusercontent.com/EduardoMartins-Dev/ambiente-check/main/coleta_ambiente.ps1" | iex
```

> Não é necessário baixar nenhum arquivo. O script roda direto na memória.

---

## O que o script faz

1. Pergunta o tipo de máquina (**Servidor** ou **Estação**)
2. Se servidor, pergunta a **faixa de usuários** para aplicar os requisitos corretos
3. Coleta informações do sistema
4. Valida cada item contra os requisitos mínimos
5. Exibe o resultado no console com `[OK]`, `[AVS]` ou `[NOK]`
6. Gera um **briefing pronto** para colar no ticket
7. Salva relatório completo em `C:\Temp\Moby_Check\`

---

## O que é verificado

| Item | Servidor | Estação |
|------|----------|---------|
| Sistema Operacional | Windows Server 2019 / Windows 10-11 | Windows 10 / 11 |
| Processador (cores) | 4 a 18 núcleos (por faixa) | — |
| Processador (clock) | ≥ 2.0 GHz | ≥ 2.0 GHz |
| Memória RAM | 8 a 64 GB (por faixa) | ≥ 4 GB |
| Disco livre | ≥ 1 TB | ≥ 4 GB |
| Resolução | 1280x600 | 1280x600 |
| .NET Framework | 2.0, 3.5, 4.0, 4.5 | 2.0, 3.5, 4.5, 4.6 |
| Navegador | Chrome ≥ 83 ou IE 11 | Chrome ≥ 83 ou IE 11 |
| Firebird | Detecta versão instalada | — |
| Banco de Dados FCerta | Localiza e mede `alterdb.ib` | — |
| Banco de Imagem FCerta | Localiza e mede `alterim.ib` | — |
| Rede (link) | ≥ 100 Mbps | ≥ 10 Mbps |
| Internet (download real) | 5 a 30 Mbps (por faixa) | ≥ 10 Mbps |
| Ping / Latência | ≤ 150ms | ≤ 150ms |
| Perda de pacotes | 0% | 0% |

---

## Faixas de usuários (Servidor)

| Faixa | Cores | RAM | Disco | Internet |
|-------|-------|-----|-------|----------|
| Até 10 | 4 | 8 GB | 1 TB | 5 Mbps |
| 11 – 20 | 6 | 16 GB | 1 TB | 10 Mbps |
| 21 – 40 | 12 | 32 GB | 1 TB | 15 Mbps |
| 41 – 60 | 18 | 64 GB | 1 TB | 30 Mbps |
| Acima de 60 | Análise técnica Fagron Tech | | | |

---

## Relatório gerado

O script salva automaticamente um `.txt` em:

```
C:\Temp\Moby_Check\Check_SERVIDOR_20260327_091814.txt
```

O arquivo contém hardware, ambiente FCerta, diagnóstico de rede (ping, velocidade, tracert) e o briefing formatado para o ticket.

---

## Requisitos para rodar

- Windows PowerShell 5.1 ou superior
- Conexão com a internet (para o teste de velocidade e download do script)
- Sem necessidade de instalar nada na máquina do cliente
