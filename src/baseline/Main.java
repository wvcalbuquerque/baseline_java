package baseline;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Locale;
import java.util.Scanner;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Main {

	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException {
		// TODO Auto-generated method stub
		
		//Linha atual.
		int linhaAtual = 0;
		
		//Células de funcionalidades		
		Cell idFuncionalidade, numFuncionalidade, nomeFuncionalidade, impactoFuncionalidade, descricaoManutencao;
		
		//Células de funções		
		Cell tipoFuncao, impactoFuncao, nomeFuncao, idFuncao, tamanhoFuncao, obsFuncao;
		
		//Array de Objetos
		ArrayList<Objeto> listaObjeto = new ArrayList<Objeto>();
		
		Scanner s = new Scanner(System.in);
		System.out.println("Informe o diretório: ");
		String caminhoDiretorio = s.next();		
		System.out.println("Caminho: " + caminhoDiretorio);
		
		File diretorio = new File(caminhoDiretorio);
		File fList[] = diretorio.listFiles();
		File planilha = new File(diretorio + "/" + fList[0].getName());			
		System.out.println("Arquivo: " + planilha);
			
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("br", "BR"));		
		Workbook pastaTrabalhoEntrada;
		pastaTrabalhoEntrada = Workbook.getWorkbook(planilha, wbSettings);
		//System.out.println("Número de abas: " + pastaTrabalhoEntrada.getNumberOfSheets());
		
		//Laço para pecorrer as abas.
		for (int i = 3; i < pastaTrabalhoEntrada.getNumberOfSheets(); i++) {
			Sheet abaOS = pastaTrabalhoEntrada.getSheet(i);
			//System.out.println(abaOS.getName());
			
			//Laço para funcionalidades.
			for (int j = 5; j < abaOS.getRows(); j++) {
				idFuncionalidade = abaOS.getCell(1,j);
				numFuncionalidade = abaOS.getCell(2,j);
				nomeFuncionalidade = abaOS.getCell(3,j);
				impactoFuncionalidade = abaOS.getCell(4,j);
				descricaoManutencao = abaOS.getCell(5,j);

				if (idFuncionalidade.getContents().equals("Expandir/Contrair")) {
					linhaAtual = j;	
					break;
				} else {						
					listaObjeto.add(new Objeto(abaOS.getName(), idFuncionalidade.getContents(), numFuncionalidade.getContents(), nomeFuncionalidade.getContents(), impactoFuncionalidade.getContents(), descricaoManutencao.getContents(), ""));							
				}
			}
			
			//Laço para funções.
			for (int k = linhaAtual+4; k < abaOS.getRows(); k++) {
				tipoFuncao = abaOS.getCell(1,k);
				impactoFuncao = abaOS.getCell(2,k);
				nomeFuncao = abaOS.getCell(3,k);
				idFuncao = abaOS.getCell(4,k);
				tamanhoFuncao = abaOS.getCell(5,k);
				obsFuncao = abaOS.getCell(6,k);
				
				if (tipoFuncao.getContents().equals("Expandir/Contrair")) {
					break;
				} else {						
					listaObjeto.add(new Objeto(abaOS.getName(), tipoFuncao.getContents(), impactoFuncao.getContents(), nomeFuncao.getContents(), idFuncao.getContents(), tamanhoFuncao.getContents(), obsFuncao.getContents()));							
				}
			}
		}		

		File file = new File(caminhoDiretorio + "/" + "baseline.xls");
		WritableWorkbook pastaTrabalhoSaida = Workbook.createWorkbook(file, wbSettings);
		pastaTrabalhoSaida.createSheet("Baseline", 0);
		WritableSheet abaSaida = pastaTrabalhoSaida.getSheet(0);
		
		Label valor = new Label (0,0, "os");
		abaSaida.addCell(valor);
		valor = new Label (1,0, "atributo1");
		abaSaida.addCell(valor);
		valor = new Label (2, 0, "atributo2");
		abaSaida.addCell(valor);
		valor = new Label (3, 0, "atributo3");
		abaSaida.addCell(valor);
		valor = new Label (4, 0, "atributo4");
		abaSaida.addCell(valor);
		valor = new Label (5, 0, "atributo5");
		abaSaida.addCell(valor);
		valor = new Label (6, 0, "atributo6");
		
		for (int l = 0; l < listaObjeto.size(); l++) {
			//System.out.println(listaObjeto.get(l).os + " | " + listaObjeto.get(l).atributo1 + " | " + listaObjeto.get(l).atributo2 + " | " + listaObjeto.get(l).atributo3 + " | " + listaObjeto.get(l).atributo4 + " | " + listaObjeto.get(l).atributo5 + " | " + listaObjeto.get(l).atributo6);
			valor = new Label (0, l+1, listaObjeto.get(l).os);
			abaSaida.addCell(valor);
			valor = new Label (1, l+1, listaObjeto.get(l).atributo1);
			abaSaida.addCell(valor);
			valor = new Label (2, l+1, listaObjeto.get(l).atributo2);
			abaSaida.addCell(valor);
			valor = new Label (3, l+1, listaObjeto.get(l).atributo3);
			abaSaida.addCell(valor);
			valor = new Label (4, l+1, listaObjeto.get(l).atributo4);
			abaSaida.addCell(valor);
			valor = new Label (5, l+1, listaObjeto.get(l).atributo5);
			abaSaida.addCell(valor);
			valor = new Label (6, l+1, listaObjeto.get(l).atributo6);
			abaSaida.addCell(valor);
		}
		
		pastaTrabalhoSaida.write();
		pastaTrabalhoSaida.close();
		
		System.out.println("Arquivo gerado: " + caminhoDiretorio + "/" + "baseline.xls");

	}

}
