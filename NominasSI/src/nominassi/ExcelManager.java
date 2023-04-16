package nominassi;

import com.sun.org.apache.xml.internal.serialize.OutputFormat;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import jdk.internal.org.xml.sax.SAXException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import controlador.Trabajador;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import org.w3c.dom.Text;

/**
 *
 * @author David
 */
public class ExcelManager {
    
    //Ubicacion excel
    // RUTA VALENTIN private String localizacionExcel ="C:/Users/valen/Documents/git/Practica_SI/NominasSI/src/resources/SistemasInformacionII.xlsx";
    //RUTA DAVID private String localizacionExcel ="C:/Users/w10/Documents/GitHub/Practica_SI/NominasSI/src/resources/SistemasInformacionII.xlsx";
    private final String localizacionExcel ="C:/Users/Torre/Documents/GitHub/Practica_SI/NominasSI/src/resources/SistemasInformacionII.xlsx";
    
    //Datos de las hojas del excel
    
    private List<Trabajador> trabajadoresHoja1= new ArrayList<>();
    
    private final Map<String, String> categoria_Complementos= new HashMap<>();
    private final Map<String, String> categoria_SalarioBase=new HashMap<>();
    
    private final Map<Float, Float> trienios= new HashMap<>();
    
    private final Map<Float, Float> retencion = new HashMap<>();
    
    private final Map<String, Float> datosEmpresa = new HashMap<>(); 
    
    //
    
    private static List<Character> letras = new ArrayList<Character>();
    private Trabajador trabajadorAux= new Trabajador();

    
    
    //Metodos para guardar la hojas del excel
    
        
    public void mapearHoja1() throws IOException, ParseException{
        
        List<String> paisOrigen = this.obtenerColumnasDatos(localizacionExcel, "Pais Origen Cuenta Bancaria", 0);
        List<String> codigoCuenta = this.obtenerColumnasDatos(localizacionExcel, "CodigoCuenta", 0);
        List<String> iban = this.obtenerColumnasDatos(localizacionExcel, "IBAN", 0);
        List<String> email = this.obtenerColumnasDatos(localizacionExcel, "Email", 0);
        List<String> fechaAltaEmpresa = this.obtenerColumnasDatos(localizacionExcel, "FechaAltaEmpresa", 0);
        List<String> cifEmpresa = this.obtenerColumnasDatos(localizacionExcel, "Cif empresa", 0);
        List<String> nombreEmpresa = this.obtenerColumnasDatos(localizacionExcel, "Nombre empresa", 0);
        List<String> categoria = this.obtenerColumnasDatos(localizacionExcel, "Categoria", 0);
        List<String> apellido1 = this.obtenerColumnasDatos(localizacionExcel, "Apellido1", 0);
        List<String> apellido2 = this.obtenerColumnasDatos(localizacionExcel, "Apellido2", 0);
        List<String> nombre = this.obtenerColumnasDatos(localizacionExcel, "Nombre", 0);
        List<String> dnis = this.obtenerColumnasDatos(localizacionExcel, "NIF/NIE", 0);
        List<String> prorrata = this.obtenerColumnasDatos(localizacionExcel, "ProrrataExtra", 0);
        List<String> fechaBajaLaboral = this.obtenerColumnasDatos(localizacionExcel, "FechaBajaLaboral", 0);
        List<String> fechaAltaLaboral = this.obtenerColumnasDatos(localizacionExcel, "FechaAltaLaboral", 0);

        int contador=2;

        for(int i=0; i<codigoCuenta.size(); i++){
            
            if(!codigoCuenta.get(i).equals("")){
               
                //System.out.println("El dni de la fila es "+" ---- "+fechaAltaLaboral.get(i)+" ----- APE "+apellido1.get(i)+" y el nombre es "+nombre.get(i));                
                //System.out.println(fechaAltaEmpresa.get(i));
                
                //Preparamos las fechas

                SimpleDateFormat formatoFecha = new SimpleDateFormat("dd-MMM-yyyy", new Locale("es", "ES"));
                SimpleDateFormat formatoFechaNuevo = new SimpleDateFormat("dd-MM-yyyy");
 
                /* Para cuando generemos las fechas  -- gestionar para cuando no hay fecha y se le pasa ""
                Date fecha_fechaAltaEmpresa = formatoFecha.parse(fechaAltaEmpresa.get(i));
                Date fecha_fechaBajaLaboral = formatoFecha.parse(fechaBajaLaboral.get(i));
                Date fecha_fechaAltaLaboral = formatoFecha.parse(fechaAltaLaboral.get(i));
 
                */

                Trabajador aux = new Trabajador(this.obtenerNumFila(localizacionExcel, codigoCuenta.get(i)), codigoCuenta.get(i)
                        , iban.get(i)
                        , email.get(i)
                        , null
                        , cifEmpresa.get(i)
                        , categoria.get(i)
                        , apellido1.get(i)
                        , apellido2.get(i)
                        , nombre.get(i)
                        , dnis.get(i)
                        , null
                        , null); 

                trabajadoresHoja1.add(aux);
     
            }

        }
        //System.out.println("FINAL "+trabajadoresHoja1.get(0).getNombre()+" locura longitud "+ trabajadoresHoja1.size());

    }
    
    
    public void mapearHoja2() throws IOException{
        
        List<String> categoria=this.obtenerColumnasDatos(localizacionExcel, "Categoria", 1);
        List<String> complementos=this.obtenerColumnasDatos(localizacionExcel, "Complementos", 1);
        List<String> salarioBase=this.obtenerColumnasDatos(localizacionExcel, "Salario Base", 1);
  
        for(int i=0; i<categoria.size(); i++){
            categoria_Complementos.put(categoria.get(i), complementos.get(i));    
            categoria_SalarioBase.put(categoria.get(i), salarioBase.get(i));
        }

        
        /*
        for (Map.Entry<String, String> entry : categoria_SalarioBase.entrySet()) {
            System.out.println(entry.getKey() + entry.getValue());
        }  */            
    }
    
    
    public void mapearHoja3() throws IOException{
       
        List<String> numTrienios = this.obtenerColumnasDatos(localizacionExcel, "Número de trienios", 2);
        List<String> impBruto = this.obtenerColumnasDatos(localizacionExcel, "Importe bruto", 2);

        for(int i=0; i<numTrienios.size(); i++){
            trienios.put(Float.parseFloat(numTrienios.get(i)), Float.parseFloat(impBruto.get(i)));
        }
        
    }
    
    public void mapearHoja4() throws IOException{
        
        List<String> brutoAnual = this.obtenerColumnasDatos(localizacionExcel, "Bruto anual", 3);
        List<String> columnaRetencion = this.obtenerColumnasDatos(localizacionExcel, "Retención", 3);
        
        for(int i=0; i<brutoAnual.size(); i++ ){
            retencion.put(Float.parseFloat(brutoAnual.get(i)), Float.parseFloat(columnaRetencion.get(i)));
        }
        
        List<String> accidentesTrabajo = this.obtenerColumnasDatos(localizacionExcel, "Accidentes trabajo EMPRESARIO", 3);
        List<String> valores = this.obtenerColumnasDatos(localizacionExcel, "1", 3);
        
        for(int j=0; j<accidentesTrabajo.size(); j++ ){
            datosEmpresa.put(accidentesTrabajo.get(j), Float.parseFloat(valores.get(j)));
        }
        
        System.out.println(datosEmpresa);
        
    }

   
    //Final metodos para guardar la hojas del excel
    public List<String> obtenerColumnasDatos(String localizacionExcel, String nombreColumna, int numHoja) throws FileNotFoundException, IOException {
        int contadorFilas = 1;
        int tope = 0; 
        int bloqueo = 0; 
        int filaActual = 0; 
        int celdaActual = 0;


        File archivoExcel = new File(localizacionExcel);                
        InputStream flujoEntrada = new FileInputStream(archivoExcel);
        XSSFWorkbook libroExcel = new XSSFWorkbook(flujoEntrada); 
        XSSFSheet hojaExcel = libroExcel.getSheetAt(numHoja); 

        Iterator<Row> iteradorFilas = hojaExcel.iterator(); 
        List<String> listaResultado = new ArrayList<>();

        while(iteradorFilas.hasNext()) 
        {
            XSSFRow fila = (XSSFRow) iteradorFilas.next();     
            Iterator<Cell> iteradorCeldas = fila.cellIterator();          

            while(iteradorCeldas.hasNext())
            {
                XSSFCell celda = (XSSFCell) iteradorCeldas.next();     
                if(celda.toString().equals(nombreColumna) && bloqueo == 0)
                {
                    tope = contadorFilas;
                    bloqueo = 1;
                }                                            
                if(bloqueo == 1 && filaActual == 1)
                {
                    if(fila.getCell(tope-1) != null && celdaActual == 0)
                    {
                        listaResultado.add(fila.getCell(tope-1).toString());
                        celdaActual = 1;
                    }else if(fila.getCell(tope-1) == null && celdaActual == 0)
                    {
                        listaResultado.add("");
                        celdaActual = 1;
                    }
                }
                contadorFilas++;
            }
            filaActual = 1;
            celdaActual = 0;
        }

        return listaResultado;
    }
    
    
    
    public void modificarDatos(String localizacionExcel, int numHoja, String antiguaCelda, String nuevaCelda) throws FileNotFoundException, IOException {


            File archivoExcel = new File(localizacionExcel);                
            InputStream flujoEntrada = new FileInputStream(archivoExcel);
            XSSFWorkbook libroExcel = new XSSFWorkbook(flujoEntrada); 
            XSSFSheet hojaExcel = libroExcel.getSheetAt(numHoja); 

            Iterator<Row> iteradorFilas = hojaExcel.iterator(); 
            List<String> listaResultado = new ArrayList<>();



            while(iteradorFilas.hasNext()) 
            {
                XSSFRow fila = (XSSFRow) iteradorFilas.next();     
                Iterator<Cell> iteradorCeldas = fila.cellIterator();      
                
                //System.out.println("la fila actual"+iteradorFilas);

                while(iteradorCeldas.hasNext())
                {
                    XSSFCell celda = (XSSFCell) iteradorCeldas.next();  
                    //System.out.println("la celda es: "+ celda.toString());
                    if(celda.toString().equals(antiguaCelda) )                   
                    {
                        celda.setCellValue(nuevaCelda);
                        //System.out.println("la celda es: "+ celda.toString());
                    }                                            

                }

            }
        flujoEntrada.close();
         //System.out.println("la celda es: ");
        
         try{
            FileOutputStream output_file = new FileOutputStream(new File(localizacionExcel));
            libroExcel.write(output_file);
            output_file.close(); 
            libroExcel.close();
            
         } catch (Exception e) {

            e.printStackTrace();
         }
       
    }
        
 
        
    //solo encuentra la primera aparicion    
    public List<String> obtenerFila(String localizacionExcel, String elemFila) throws FileNotFoundException, IOException{    //devuelve una lista con los elementos de una fila. La fila sera en la que se encuentre elemFila
    
        File archivoExcel = new File(localizacionExcel);                
        InputStream flujoEntrada = new FileInputStream(archivoExcel);
        XSSFWorkbook libroExcel = new XSSFWorkbook(flujoEntrada); 
        XSSFSheet hojaExcel = libroExcel.getSheetAt(0); 

        Iterator<Row> iteradorFilas = hojaExcel.iterator(); 
        List<String> listaResultado = new ArrayList<>();
        boolean encontrado=false;



        while(iteradorFilas.hasNext() && encontrado==false) 
        {
            XSSFRow fila = (XSSFRow) iteradorFilas.next(); 
            //System.out.println("NumFIla es: "+fila.getRowNum());          //PARA OBTENER EL NUMERO DE LA FILA DEL EXCEL
            Iterator<Cell> iteradorCeldas = fila.cellIterator();          

            while(iteradorCeldas.hasNext())
            {
                XSSFCell celda = (XSSFCell) iteradorCeldas.next();

                
                if(celda.toString().equals(elemFila)) {
                    encontrado=true;
                    int num=0;
                    Iterator<Cell> iteradorCeldasFila = fila.cellIterator();        //creamos nuevo iterador para la fila
                    while(iteradorCeldasFila.hasNext()) {
                        XSSFCell celdaFila = (XSSFCell) iteradorCeldasFila.next();
                        num=fila.getRowNum()+1;
                        
                        listaResultado.add(celdaFila.toString());
                        
                        System.out.println(celdaFila.toString());
                        
                        
                    }
                    listaResultado.add(""+num);
                    
                    
                }
            }

        }
        
      
        return listaResultado;
    }
    
    public List<Integer> obtenerRepetidos(String dniRepetido){
        
        int contador=0;
        List<Integer> resultado = new ArrayList<>();
        
        for(int i=0; i<trabajadoresHoja1.size(); i++){
            
            if(trabajadoresHoja1.get(i).getNifnie().equals(dniRepetido) && contador>0){
                resultado.add(i);                
            }else if(trabajadoresHoja1.get(i).getNifnie().equals(dniRepetido)){
                contador++;
            }
            
        }
        
        //Devuelve una lista con los numeros de los trabajadores con los dnis repetidos
        
        return resultado;
    }
    
    
    public List<String> obtenerFilaRepeticiones(String localizacionExcel, String elemFila, int repeticion) throws FileNotFoundException, IOException{    //devuelve una lista con los elementos de una fila. La fila sera en la que se encuentre elemFila
    
        File archivoExcel = new File(localizacionExcel);                
        InputStream flujoEntrada = new FileInputStream(archivoExcel);
        XSSFWorkbook libroExcel = new XSSFWorkbook(flujoEntrada); 
        XSSFSheet hojaExcel = libroExcel.getSheetAt(0); 

        Iterator<Row> iteradorFilas = hojaExcel.iterator(); 
        List<String> listaResultado = new ArrayList<>();
        boolean encontrado=false;
        
        int aux=repeticion-1;

        while(iteradorFilas.hasNext() && encontrado==false) 
        {
            XSSFRow fila = (XSSFRow) iteradorFilas.next(); 
            //System.out.println("NumFIla es: "+fila.getRowNum());          //PARA OBTENER EL NUMERO DE LA FILA DEL EXCEL
            Iterator<Cell> iteradorCeldas = fila.cellIterator();          

            while(iteradorCeldas.hasNext())
            {
                XSSFCell celda = (XSSFCell) iteradorCeldas.next();

                
                if(celda.toString().equals(elemFila) && aux==0) {
                    
                    
                    encontrado=true;
                    int num=0;
                    Iterator<Cell> iteradorCeldasFila = fila.cellIterator();        //creamos nuevo iterador para la fila
                    while(iteradorCeldasFila.hasNext()) {
                        XSSFCell celdaFila = (XSSFCell) iteradorCeldasFila.next();
                        num=fila.getRowNum()+1;
                        
                        listaResultado.add(celdaFila.toString());                     
                        
                    }
                    listaResultado.add(""+num);
                    
                    
                    
                    
                    
                }else if(celda.toString().equals(elemFila)){
                    aux= aux-1;
                }
            }

        }

        
      
        return listaResultado;
    }

    public Map<String, Integer> contarRepeticiones(List<String> lista){
       
        Map<String, Integer> map = new HashMap<>();
        for (String elemento : lista) {
            if (map.containsKey(elemento)) {
                map.put(elemento, map.get(elemento) + 1);
            } else {
                map.put(elemento, 1);
            }
        }
       
        
        return map;
    }
    
    public int obtenerNumFila(String localizacionExcel, String elemFila) throws FileNotFoundException, IOException{    //devuelve una lista con los elementos de una fila. La fila sera en la que se encuentre elemFila
    
        File archivoExcel = new File(localizacionExcel);                
        InputStream flujoEntrada = new FileInputStream(archivoExcel);
        XSSFWorkbook libroExcel = new XSSFWorkbook(flujoEntrada); 
        XSSFSheet hojaExcel = libroExcel.getSheetAt(0); 

        Iterator<Row> iteradorFilas = hojaExcel.iterator(); 
        boolean encontrado=false;
        int num=-1;
        
        while(iteradorFilas.hasNext() && encontrado==false) 
        {
            XSSFRow fila = (XSSFRow) iteradorFilas.next(); 
            Iterator<Cell> iteradorCeldas = fila.cellIterator();          

            while(iteradorCeldas.hasNext())
            {
                XSSFCell celda = (XSSFCell) iteradorCeldas.next();
                
                if(celda.toString().equals(elemFila)) {
                    encontrado=true;
                    num=fila.getRowNum()+1;
                }
            }
        }
        return num;
    }    
        
    public void procesarDNI() throws FileNotFoundException, IOException, ParserConfigurationException, SAXException, org.xml.sax.SAXException {
        
        // SE RELLENA LA LISTA QUE CONTIENE LAS LETRAS DE LOS DNI
        char[] listaAux = new char[]{'T', 'R', 'W', 'A', 'G', 'M', 'Y', 'F', 'P', 'D', 'X', 'B', 'N', 'J', 'Z', 'S', 'Q', 'V', 'H', 'L', 'C', 'K', 'E'};
        List<Trabajador> trabajadoresErrores= new ArrayList<>();
        for(int i=0; i<23; i++) {
            letras.add(listaAux[i]);
        }
        

        //List<String> listaDNI = this.obtenerColumnasDatos(localizacionExcel, "NIF/NIE", 0);
        List<String> listaDNI= new ArrayList<>();
        
        for(int j=0; j< trabajadoresHoja1.size(); j++){
            listaDNI.add(trabajadoresHoja1.get(j).getNifnie());
        }
        Map<String, Integer> map = contarRepeticiones(listaDNI);
        List<String> listaDNI_Repetidos = new ArrayList<>();
        
        
        
        for(int i=0; i<trabajadoresHoja1.size(); i++){       
            
            
                
                //System.out.println("el elemento "+listaDNI.get(i)+" se repite estas veces: "+map.get(listaDNI.get(i)) );

                if(map.get(listaDNI.get(i))>1 && !listaDNI_Repetidos.contains(listaDNI.get(i)) && !listaDNI.get(i).equals("")){  //comprobar que el dni se repite y que no se encuentra en la lista de "ya añadidos"

                    //System.out.println("Añadidos a XML ERRORES por repeticion: "+ listaDNI.get(i)+" "+ map.get(listaDNI.get(i)));
                    
                    List<Integer> numTrabajadoresRepetidos = obtenerRepetidos(listaDNI.get(i));                
                    
                    for(int j=0; j< numTrabajadoresRepetidos.size(); j++){ 
                       
                        trabajadoresErrores.add(trabajadoresHoja1.get(numTrabajadoresRepetidos.get(j)));
                        
                    }
                    
                    listaDNI_Repetidos.add(listaDNI.get(i));
                }else{
              
                    int comprobacion=esValidoDNI(listaDNI.get(i));

                    switch(comprobacion){

                        case 2:
                            //el error se puede subsanar -> LA LETRA ESTA MAL
                            
                            String dniArreglado = arreglarDNI(listaDNI.get(i)); 
                            this.modificarDatos(localizacionExcel, 0, listaDNI.get(i), dniArreglado);
                            trabajadoresHoja1.get(i).setNifnie(dniArreglado);
                            break;
                        case 3:
                            //el error no es subsanable -> ESTÁ MAL ESTRUCTURADO -> añadir al XML
                            
                           trabajadoresErrores.add(trabajadoresHoja1.get(i));
                           break;
                    }
                }
            
        }
       
        if(trabajadoresErrores.size()>0){
            
            try {
                this.agregarTrabajadoresAXML(trabajadoresErrores);
            } catch (TransformerException ex) {
               System.out.println(ex);
            }
            
        }
  
    }        
    
    public static int esValidoDNI(String dni) {
        
        // RETURN:
        // 1 - VALIDO
        // 2 - ERROR SUBSANABLE
        // 3 - ERROR NO SUBSANABLE
    
        int esValido = 3;
        char letra;
        int cantidad;

        if (dni.length() == 9) {   //el dni tiene longitud 9
                  
            if (estaBienEstructurado(dni)) {
                
                letra = dni.charAt(8);
                cantidad = Integer.parseInt(dni.substring(0, dni.length()-1));

                if (letra == obtenerLetraCorrectaDNI(cantidad)) {  //si la letra es la correcta, el dni es valido

                    esValido = 1;
                } else {    // si la letra no es la correcta, el dni es erroneo pero subsanable
                
                    esValido = 2;
                }
            } 
        }

        return esValido;
    }
    
    public static int esValidoNIE(String nie) {     //SIN TERMINAR
        
        // RETURN:
        // 1 - VALIDO
        // 2 - ERROR SUBSANABLE
        // 3 - ERROR NO SUBSANABLE
    
        int esValido = 3;
        char letra;
        int cantidad;

        if (nie.length() == 9) {   //el dni tiene longitud 9
                  
            if (estaBienEstructurado(nie)) {
                
                letra = nie.charAt(8);
                cantidad = Integer.parseInt(nie.substring(0, nie.length()-1));

                if (letra == obtenerLetraCorrectaDNI(cantidad)) {  //si la letra es la correcta, el dni es valido

                    esValido = 1;
                } else {    // si la letra no es la correcta, el dni es erroneo pero subsanable
                
                    esValido = 2;
                }
            } 
        }

        return esValido;
    }

    public static boolean estaBienEstructurado(String dni) {
    
        boolean estaBien = true;
        boolean encontrado = false;
        char letra;

            for (int i=0; i<9; i++) {

                if (i<8) {    // PARA LOS 8 PRIMEROS DIGITOS SE COMPRUEBA QUE SEAN NUMEROS
                    encontrado = false;
                    letra = dni.charAt(i);

                    if (letra == '1' || letra == '2'|| letra == '3' || letra == '4' || letra == '5' || letra == '6' || letra == '7' || letra == '8' || letra == '9' || letra == '0') {
                        encontrado = true;
                    }
                    if (!encontrado) {
                        estaBien = false;
                    }
                } else {   // PARA EL ULTIMO DIGITO SE COMPRUEBA QUE SEA UNA DE LAS LETRAS VALIDAS
                    encontrado = false;
                    for (int j=0; j<letras.size(); j++) {
                        if (dni.charAt(8) == letras.get(j)) {
                            encontrado = true;
                        }
                    }
                    if (!encontrado) {
                        estaBien = false;
                    }
                }
            }     
        
        return estaBien;
    }

    public static char obtenerLetraCorrectaDNI(int suma){   //se devuelve la letra correcta teniendo en cuenta el numero
        int numeroBusqueda = suma % 23;
        char caracterRetorno = letras.get(numeroBusqueda);
        return caracterRetorno;
    }
    
    public static String arreglarDNI(String dni){

        String nuevoDNI = String.copyValueOf(dni.toCharArray());
        int cantidad = Integer.parseInt(dni.substring(0, dni.length()-1));
        char letra = obtenerLetraCorrectaDNI(cantidad);     //se obtiene la letra correspondiente al numero 
        nuevoDNI = ((cantidad + "")+ letra);   

        if (nuevoDNI.length() < 9) {
            for (int i=nuevoDNI.length(); i<9; i++) {
                nuevoDNI = ('0'+nuevoDNI);
            }
        }
        
        return nuevoDNI;
    }
    
    
    
    
public void agregarTrabajadoresAXML(List<Trabajador> trabajadores) throws ParserConfigurationException, IOException, SAXException, TransformerException, org.xml.sax.SAXException {

    
        try{
        // cargamos el archivo XML existente en un objeto Document

        // RUTA DAVID String rutaXML = "C:/Users/w10/Documents/GitHub/Practica_SI/NominasSI/src/resources/Errores.xml";
        // RUTA LAPTOP String rutaXML = "C:/Users/valen/Documents/git/Practica_SI/NominasSI/src/resources/Errores.xml";
        String rutaXML = "C:/Users/Torre/Documents/GitHub/Practica_SI/NominasSI/src/resources/Errores.xml";
        


        File archivoXML = new File(rutaXML);
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        
        DocumentBuilder db = dbf.newDocumentBuilder();
       
        Document doc = db.newDocument();
        Element rootElement = doc.createElement("Trabajadores");
        doc.appendChild(rootElement);

        // obtenemos la raíz del documento existente
        Element eRaiz = doc.getDocumentElement();

        // creamos un nuevo elemento para cada trabajador
        for (int i = 0; i < trabajadores.size(); i++) {

            Element xmlTrabajador = doc.createElement("Trabajador");

            
            Attr atributoID = doc.createAttribute("id");
            atributoID.setValue(""+trabajadores.get(i).getIdTrabajador());
            xmlTrabajador.setAttributeNode(atributoID);
            
            Element nif = doc.createElement("NIF_NIE");
            nif.appendChild(doc.createTextNode(trabajadores.get(i).getNifnie()));
            xmlTrabajador.appendChild(nif);

            
            Element nombre = doc.createElement("Nombre");
            nombre.appendChild(doc.createTextNode(trabajadores.get(i).getNombre()));
            xmlTrabajador.appendChild(nombre);
            
            

            Element apellido1 = doc.createElement("PrimerApellido");
            apellido1.appendChild(doc.createTextNode(trabajadores.get(i).getApellido1()));
            xmlTrabajador.appendChild(apellido1);

            Element apellido2 = doc.createElement("SegundoApellido");
            apellido2.appendChild(doc.createTextNode(trabajadores.get(i).getApellido2()));
            xmlTrabajador.appendChild(apellido2);
            

            Element empresa = doc.createElement("Empresa");
            empresa.appendChild(doc.createTextNode(trabajadores.get(i).getEmpresa()));
            xmlTrabajador.appendChild(empresa);

            Element categoria = doc.createElement("Categoria");
            categoria.appendChild(doc.createTextNode(trabajadores.get(i).getCategoria()));
            xmlTrabajador.appendChild(categoria);

            // añadimos el elemento del trabajador a la raíz del documento
            eRaiz.appendChild(xmlTrabajador);
        }

        // actualizamos el archivo XML
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes"); // configuramos la propiedad para que se escriba en varias líneas
        DOMSource source = new DOMSource(doc);
        StreamResult result = new StreamResult(archivoXML);
        transformer.transform(source, result);

        }catch(Exception e){
            e.printStackTrace();
        }
    }
    
}