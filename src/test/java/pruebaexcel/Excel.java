package pruebaexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	// XSSFWorkbook -> .xlsx
	private XSSFWorkbook libroExcel = null;
	private XSSFSheet hojaLibro;
	int ultimaFila;
	// HSSF -> .xls
	public Excel(String strNombreHoja) {
		libroExcel = new XSSFWorkbook();
		hojaLibro = libroExcel.createSheet(strNombreHoja);
		ultimaFila = hojaLibro.getLastRowNum();
	}
	
	public Excel(String strRutaArchivo, String strHoja) {
		File file = new File(strRutaArchivo);
		
		try {
			FileInputStream fileInput = new FileInputStream(file);
			libroExcel = new XSSFWorkbook(fileInput);
		} catch (IOException e) {
			e.printStackTrace();
		}
		hojaLibro = libroExcel.getSheet(strHoja);
	}
	
	public void leerTodosLosDatos() {
		int intUltimaFila = hojaLibro.getLastRowNum();
		for (int i = 0; i < intUltimaFila; i++) {
			Row fila = hojaLibro.getRow(i);
			int intUltimaColumna = fila.getLastCellNum();
			for (int j = 0; j < intUltimaColumna; j++) {
				Cell celda = fila.getCell(j);
				if (celda.getCellTypeEnum() == CellType.STRING) {
                    System.out.print("|" + celda.getStringCellValue());
                } else if (celda.getCellTypeEnum() == CellType.NUMERIC) {
                    System.out.print("|"+celda.getNumericCellValue());
                }
				
		
				
			}
			
			System.out.print("|");
			System.out.println("");
		}
	}
	public void leerCeldaEsp(int intFila, int intColumna) {
		Row fila = hojaLibro.getRow(intFila-1);
		Cell celda = fila.createCell(intColumna-1);
		 System.out.print("|" + celda.getStringCellValue());
		
	}
	public void escribirExcelEnFilaYColumnaNueva(String strValor) {
		
		Row fila;
		if(ultimaFila == 0) {
			fila = hojaLibro.createRow(ultimaFila);
			ultimaFila++;
		}
		else {
			fila = hojaLibro.createRow(ultimaFila++);
		}
		Cell celda = fila.createCell(0); 
		celda.setCellValue(strValor);
		
	}
	public void escribirEnFilaYColumnaEspecifica(String strValor, 
			int intFila, int intColumna) {
		Row fila = hojaLibro.createRow(intFila-1);
		Cell celda = fila.createCell(intColumna-1);
		celda.setCellValue(strValor);
	}
	
	public void exportarArchivo(String strRuta) {
		File file = new File(strRuta);
		try {
			FileOutputStream fileOutput = new FileOutputStream(file);
			libroExcel.write(fileOutput);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
