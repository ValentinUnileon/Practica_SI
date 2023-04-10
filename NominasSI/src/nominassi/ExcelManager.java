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
import org.w3c.dom.Text;

/**
 *
 * @author David
 */
public class ExcelManager {


    private static List<Character> letras = new ArrayList<Character>();
    private Trabajador trabajadorAux= new Trabajador();

  
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
                        
                        
                    }
                    listaResultado.add(""+num);
                    
                    
                }
            }

        }

        
      
        return listaResultado;
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
        
    public void procesarDNI(String localizacionExcel) throws FileNotFoundException, IOException, ParserConfigurationException, SAXException, org.xml.sax.SAXException {
        
        // SE RELLENA LA LISTA QUE CONTIENE LAS LETRAS DE LOS DNI
        char[] listaAux = new char[]{'T', 'R', 'W', 'A', 'G', 'M', 'Y', 'F', 'P', 'D', 'X', 'B', 'N', 'J', 'Z', 'S', 'Q', 'V', 'H', 'L', 'C', 'K', 'E'};
        List<Trabajador> trabajadoresErrores= new ArrayList<>();
        for(int i=0; i<23; i++) {
            letras.add(listaAux[i]);
        }
        

        List<String> listaDNI = this.obtenerColumnasDatos(localizacionExcel, "NIF/NIE", 0);
        Map<String, Integer> map = contarRepeticiones(listaDNI);
        List<String> listaDNI_Repetidos = new ArrayList<>();
        
        
        
        for(int i=0; i<listaDNI.size(); i++){       
            
            if(!listaDNI.get(i).equals("")){
                
                //System.out.println("el elemento "+listaDNI.get(i)+" se repite estas veces: "+map.get(listaDNI.get(i)) );
                
              
                
                if(map.get(listaDNI.get(i))>1 && !listaDNI_Repetidos.contains(listaDNI.get(i))){  //comprobar que el dni se repite y que no se encuentra en la lista de "ya añadidos"

                    
                    //System.out.println("Añadidos a XML ERRORES por repeticion: "+ listaDNI.get(i)+" "+ map.get(listaDNI.get(i)));
                     
                    
                    for(int j=2; j< map.get(listaDNI.get(i))+1; j++){
                        
                        
                        
                        List<String> filaTrabajador= this.obtenerFilaRepeticiones(localizacionExcel, listaDNI.get(i), j);
                        
                        

                       
                        
                        Trabajador trabajadorProvisional1 = trabajadorAux.rellenarTrabajadorExcel(filaTrabajador);
                        //System.out.println(trabajadorProvisional1.getNombre()+" repetido");

                        trabajadoresErrores.add(trabajadorProvisional1);

                        
                    }
                    
                    listaDNI_Repetidos.add(listaDNI.get(i));
                }else{
                                    
                int comprobacion=esValidoDNI(listaDNI.get(i));
                
                
                
                
                switch(comprobacion){

                    case 2:
                        //el error se puede subsanar -> LA LETRA ESTA MAL
                        String dniArreglado = arreglarDNI(listaDNI.get(i));  //DNI CON LA LETRA CORRECTA
                        this.modificarDatos(localizacionExcel, 0, listaDNI.get(i), dniArreglado);
                        //.out.println("El dni: "+listaDNI.get(i)+" ha sido reemplazado por "+dniArreglado);
                        break;
                    case 3:
                        //el error no es subsanable -> ESTÁ MAL ESTRUCTURADO -> añadir al XML
                        
                        
                       

                        
                        List<String> filaTrabajador= this.obtenerFila(localizacionExcel, listaDNI.get(i));

                        //System.out.println("error del dni al xml " + listaDNI.get(i));

                        
                        
                        try{
                        Trabajador trabajadorProvisional = trabajadorAux.rellenarTrabajadorExcel(filaTrabajador);
                       

                        trabajadoresErrores.add(trabajadorProvisional);
                        }catch(Exception e){
                            //System.out.println("LA LONGITUD ES "+ filaTrabajador.size());
                            e.printStackTrace();
                        }
                        
                       break;

                               
                }
                    
                    
                    
                    
                    
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

        String rutaXML = "C:/Users/w10/Documents/GitHub/Practica_SI/NominasSI/src/resources/Errores.xml";


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