# Ajuste PGCC DeRE

Ferramenta de uma página para preparar o plano de contas (evento D-1011 da DeRE) exportado do DeRE ("Exportar PGCC").

- Arraste o arquivo (.csv ou .xlsx) na página.
- Sai na hora um .zip com: CSV para "Importar CSV (150k)", Excel e PDF de conferência.
- Regras: sem acento/Ç/º, só letras, números e espaço (mantém "( - )"), "&" vira "E", e nenhum caractere proibido do manual DeRE 1.6 (" ' -- # ? /* e seus escapes).

Tudo roda no navegador de quem usa: **o arquivo não é enviado para nenhum servidor**. A página não tem dados embutidos (só um exemplo fictício).

Bibliotecas embutidas (licenças nos próprios arquivos): SheetJS 0.18.5 (Apache-2.0), JSZip 3.10.1 (MIT), jsPDF 2.5.1 (MIT), jspdf-autotable 3.8.2 (MIT).
