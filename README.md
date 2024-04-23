### Garza Inteligência Financeira
## Projeto Convergência


### *Pedro Henrique Campos Moreira*
---

## 1. Instale o Visual Studio Code (Vscode)

  - **Entre na Microsoft Store**  
  - **Pesquise Visual studio Code**
  - **Instale a IDE**

## 2. Instale o Python (Versão mais recente e compatível com o seu S.O.)

   ```shell
   https://www.python.org/downloads/
   ```

## 3. Clone este repositório:

  - **Copie o arquivo projeto27.py**  
  - **Abra o Vscode**
  - **Crie um novo arquivo.py (file)**

## 4. Bibliotecas - Importações

**Bibliotecas devem ser baixadas no terminal.**

   ```shell
  pip install openpyxl
   ```

   ```shell
   pip install matplotlib
   ```

   ```shell
 pip install PyPDF2
   ```

   ```shell
  pip install pywin32
   ```

## 5. Preparo

  - **Cole o código no novo arquivo criado**  
  - **Se necessário, adicione o arquivo clientes.xlxs em sua pasta**
  - **Ter os arquivos pdf's e excel baixados em sua máquina**

## 6. Execução

 1. **Compile o código**  
 2. **Escolha sua opção**
    - (Sim) Adcione o cliente
 3. **Nome do cliente que deseja consultar**
 4. **Selecione a planilha e o arquivo do cliente desejados**
 5. Verifique seu e-mail**
   
## 7. Observações

```
    if tipo_cliente == "Balanceado":
        porcentagens_excel = [46.85, 7.5, 22.1, 0, 8.04, 9.6, 6]
    elif tipo_cliente == "Moderado":
        porcentagens_excel = [62.05, 5, 20.23, 0, 4.32, 4.8, 3.6]
    elif tipo_cliente == "Conservador":
        porcentagens_excel = [88.5, 2.5, 5.8, 0, 3.2, 0, 0]
    elif tipo_cliente == "Arrojado":
        porcentagens_excel = [34.66, 8.75, 22.27, 0, 9.02, 13.3, 9.5]
    else:
        porcentagens_excel = [42.9, 5, 20, 0, 12.6, 12, 7.5]
```
 - Modificar o codigo acima pode alterar os resultados do perfil do cliente de acordo com a planilha, podendo adicionar novos perfis e porcentagens 

##
```
def enviar_email(cliente, perfil, texto, path_combined_image):
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.To = 'pedro.moreira@garzaif.com.br'
    email.Subject = f"Convergência - {cliente}"
    email.HTMLBody = texto + "<img src='cid:imagem_cid'>"
    attachment = email.Attachments.Add(path_combined_image)
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001E", "imagem_cid")
    email.Send()
    outlook.Quit()
  ```
 - Deve-se alterar o codigo acima para receber o email de forma correta

 ## 8. Recomandações
  - Em caso de erro , fechar todas as jenals e iniciar novamente o codigo
  - Todos os arquivos pdf devem estar no mesmo padrão, exibindo apenas o grafico redondo e nenhum outro
  - Todos os arquivos pdf que possuem muitas informaões sua formatação deve ser em apenas uma coluna (informações dos ativos devem estar na mesma pagina)
---
