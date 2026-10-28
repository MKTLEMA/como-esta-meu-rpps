# Como está meu RPPS? | LEMA Edu

Consulta de porte, índice de situação previdenciária (ISP) e certificação exigida por cargo no seu RPPS.

- Site: https://mktlema.github.io/como-esta-meu-rpps/
- Arquivo para baixar e abrir (funciona offline, no celular ou no computador): [index.html](https://mktlema.github.io/como-esta-meu-rpps/index.html) (clique com o botão direito e salve como).

## Como se atualiza sozinho
Todo dia, às 06h (Brasília), o GitHub Actions roda `atualizar_dados.py`: lê a página oficial do ISP-RPPS no gov.br, baixa o resultado final mais recente e o do ano anterior e troca os dados dentro do `index.html` quando houver mudança. O site republica em seguida. O mesmo job vigia a página do Manual de Certificação: se sair versão maior que a usada (hoje 2.0), o job falha e o GitHub avisa por e-mail para revisar as regras.

## Fontes
Ministério da Previdência Social (SPREV): ISP-RPPS (resultado final), Portaria MPS 1.499/2024 e Manual de Certificação dos Profissionais dos RPPS, versão 2.0 (Portaria SRPC/MPS 607/2026). Informativo: confira sempre a norma.
