# WordsReplace

Este projeto é um estudo que manipula arquivos do tipo DOCX do Microsoft Word 

___

* Lembrando que você deve criar dois arquivos Word, garantir acesso ao caminho do arquivo de leitura e gravação do diretório e que o arquivo que irá está no diretório pathDocDefault deverá conter a palavra ou termo que irá ser substituido.


O primeiro método substitui palavras que estão no arquivo, gerando assim um novo arquivo com a apalavra substituída


Request exemplo para substituir as palavras


```js
{
  "pathDocDefault": "C:\\\\Teste\\doc_default_pag1.docx",
  "pathNewDoc": "C:\\\\Teste\\doc_pag1.docx",
  "oldWord": "%palavra_antiga%",
  "newWord": "teste"
}
```


O segundo método serve para juntar o novo arquivo criado anteriormente com uma novo arquivo, gerando assim um novo arquivo com os dois contaúdos, um em cada página. 


Request exemplo para juntar dois documentos (um em cada página)


```js
{
	"filePath1": "C:\\\\Teste\\doc_pag1.docx",
	"filePath2": "C:\\\\Teste\\doc_default_pag2.docx",
	"mergedFilePath": "C:\\\\Teste\\new_doc_merge.docx"
}
```