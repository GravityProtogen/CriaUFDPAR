from docx import Document
import docxedit
path = 'c:/Users/infra/Desktop/Pedro/Coding/Doc reader/ExtrairNome/InserirNome/TceuTemplate.docx'
document = Document(path)
# Inputs
nome = input("Nome do estagiario: ")
matricula = input("Matricula do estagiario: ")
identidade = input("Identidade do estagiario: ")
orgao = input("Orgao emitente da identidade: ")
uf = input("UF da identidade: ")
cpf = input("CPF do estagiario: ")
data = input("Data de nascimento do estagiario: ")
naturalidade = input("Naturalidade do estagiario: ")
sexo = input("Sexo do estagiario: ")



# Replace
docxedit.replace_string(document, old_string='NomeEstagiario', new_string=nome)
docxedit.replace_string(document, old_string='MatriculaEstagiario', new_string=matricula)
docxedit.replace_string(document, old_string='IdEstagiario', new_string=identidade)
docxedit.replace_string(document, old_string='OrgaoEstagiario', new_string=orgao)
docxedit.replace_string(document, old_string='UFEstagiario', new_string=uf)
docxedit.replace_string(document, old_string='CPFEstagiario', new_string=cpf)
docxedit.replace_string(document, old_string='DataEstagiario', new_string=data)
docxedit.replace_string(document, old_string='NaturalidadeEstagiario', new_string=naturalidade)
if sexo.lower() == 'm':
    docxedit.replace_string(document, old_string='MascEstagiario', new_string="X")
    docxedit.replace_string(document, old_string='FemEstagiario', new_string=" ")
elif sexo.lower() == 'f':
    docxedit.replace_string(document, old_string='FemEstagiario', new_string="X")
    docxedit.replace_string(document, old_string='MascEstagiario', new_string=" ")


document.save('tceufdpar-prenchido.docx')

