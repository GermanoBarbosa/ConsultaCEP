# ConsultaCEP

Um projeto em **Visual Basic 6 (VB6)** para consultar endereços completos a partir de um CEP, utilizando a **BrasilAPI**. A consulta é feita via requisição HTTP, com resultados retornados em formato JSON. O projeto oferece uma interface simples onde o usuário insere o CEP e visualiza informações como Rua, Bairro, Cidade e Estado.

## Funcionalidades

- **Consulta de CEP**: Insira um CEP válido para obter os dados do endereço via BrasilAPI.
- **Exibição em tempo real**: Exibe Rua, Bairro, Cidade e Estado diretamente na interface.
- **Interface intuitiva**: Utiliza controles nativos do VB6 (TextBox, Label, Command semperButton).

## Tecnologias Utilizadas

- **Visual Basic 6 (VB6)**: Plataforma de desenvolvimento.
- **BrasilAPI**: API para consulta de CEPs.
- **VBA-JSON**: Biblioteca para parse de respostas JSON.

## Pré-requisitos

- Visual Basic 6 instalado.
- Biblioteca **Microsoft XML, v6.0** (ou similar) para requisições HTTP.
- Classe **JsonConverter.cls** para parse de JSON.

## Como Rodar

1. **Clonar o repositório**:
   ```bash
   git clone https://github.com/GermanoBarbosa/ConsultaCEP.git
   ```

2. **Adicionar a classe JSON**:
   - Baixe a classe `JsonConverter.cls` (disponível em repositórios como [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)).
   - No VB6, vá em **Project > Add File** e inclua a classe no projeto.

3. **Configurar referência HTTP**:
   - No VB6, acesse **Project > References**.
   - Marque **Microsoft XML, v6.0** (ou versão compatível).

4. **Executar o projeto**:
   - Abra o arquivo `frmConsultaCEP.frm` no VB6.
   - Compile e execute o projeto.
   - Insira um CEP válido (ex.: `01001-000`) e clique em **Consultar CEP**.

## Exemplo de Uso

1. Abra o aplicativo no VB6.
2. Digite um CEP no campo de texto (ex.: `01001-000`).
3. Clique em **Consultar CEP**.
4. Visualize os dados de Rua, Bairro, Cidade e Estado exibidos na interface.

## Licença

Este projeto está sob a licença **MIT**. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.