import pandas as pd

# Dados de exemplo para preencher o modelo
data = {
    'Nome': ['John Doe', 'Jane Smith', 'Bob Johnson'],
    'Email': ['johndoe@example.com', 'janesmith@example.com', 'bobjohnson@example.com'],
    'Assunto': ['Hello', 'Greetings', 'Important Announcement'],
    'Mensagem': ['Hi John, How are you?', 'Dear Jane, Have a great day!', 'Hello Bob, Please read the attached document.'],
    'Anexo': ['C:\\path\\to\\attachment1.pdf', 'C:\\path\\to\\attachment2.docx', '']
}

# Criar DataFrame a partir dos dados
df = pd.DataFrame(data)

# Salvar DataFrame em um arquivo Excel
df.to_excel('modelo_emails2.xlsx', index=False)
