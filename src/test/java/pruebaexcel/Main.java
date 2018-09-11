package pruebaexcel;

public class Main {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Excel excel = new Excel("Prueba");
		excel.escribirExcelEnFilaYColumnaNueva("Juan");
		excel.escribirExcelEnFilaYColumnaNueva("Pedro");
		excel.escribirExcelEnFilaYColumnaNueva("Metro");
		excel.escribirEnFilaYColumnaEspecifica("Fila espec", 10, 10);
		excel.exportarArchivo("C:\\prueba23052018.xlsx");
		
//		Excel excelLectura = new Excel("C:\\ruta.xlsx", "Cualquier nombre");
//		excelLectura.leerTodosLosDatos();
		Excel lectura2 = new Excel ("C:\\prueba23052018.xlsx", "Prueba");
		lectura2.leerCeldaEsp(1,1);
	}

}
