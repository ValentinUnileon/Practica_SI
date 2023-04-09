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


        File archivoExcel = new File(localizacionExcel);                //ponerTrycatch
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


            File archivoExcel = new File(localizacionExcel);                //ponerTrycatch
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
                        System.out.println("la celda es: "+ celda.toString());

                    }                                            


                }

            }
        flujoEntrada.close();
         System.out.println("la celda es: ");
        
         try{
             FileOutputStream output_file = new FileOutputStream(new File(localizacionExcel));
                     libroExcel.write(output_file);
        
        output_file.close(); 
        
        libroExcel.close();
         } catch (Exception e) {
    // Manejo de la excepción
            e.printStackTrace();
         }
        
         
       

        




        }
        
        
    //solo encuentra la primera aparicion    
    public List<String> obtenerFila(String localizacionExcel, String elemFila) throws FileNotFoundException, IOException{    //devuelve una lista con los elementos de una fila. La fila sera en la que se encuentre elemFila
    
        File archivoExcel = new File(localizacionExcel);                //ponerTrycatch
        InputStream flujoEntrada = new FileInputStream(archivoExcel);
        XSSFWorkbook libroExcel = new XSSFWorkbook(flujoEntrada); 
        XSSFSheet hojaExcel = libroExcel.getSheetAt(0); 

        Iterator<Row> iteradorFilas = hojaExcel.iterator(); 
        List<String> listaResultado = new ArrayList<>();
        boolean encontrado=false;



        while(iteradorFilas.hasNext() && encontrado==false) 
        {
            XSSFRow fila = (XSSFRow) iteradorFilas.next(); 
            //System.out.println("NumFIla esss: "+fila.getRowNum());          //PARA OBTENER EL NUMERO DE LA FILA DEL EXCEL
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
                        num=fila.getRowNum();
                        
                        listaResultado.add(celdaFila.toString());
                        
                        
                    }
                    listaResultado.add(""+num);
                    
                    
                }
            }

        }

        
      
        return listaResultado;
    }
    
    
        public List<String> obtenerFilaRepeticiones(String localizacionExcel, String elemFila, int repeticion) throws FileNotFoundException, IOException{    //devuelve una lista con los elementos de una fila. La fila sera en la que se encuentre elemFila
    
        File archivoExcel = new File(localizacionExcel);                //ponerTrycatch
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
            //System.out.println("NumFIla esss: "+fila.getRowNum());          //PARA OBTENER EL NUMERO DE LA FILA DEL EXCEL
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
                        num=fila.getRowNum();
                        
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
        
        for(int i=0; i<listaDNI.size(); i++){       //hacer metodo para encontrar elems con mas de una aparicion (repetidos)
            
            if(!listaDNI.get(i).equals("")){
                
                System.out.println("el elemento "+listaDNI.get(i)+" se repite estas veces: "+map.get(listaDNI.get(i)) );
                
              
                
                if(map.get(listaDNI.get(i))>1 && !listaDNI_Repetidos.contains(listaDNI.get(i))){  //comprobar que el dni se repite y que no se encuentra en la lista de "ya añadidos"
                    
                    
                    //introducir todos los trabajdores repetido al XML de errores menos el primero********
                    //hacer metodo  obtenerFilaRepeticiones(String localizacionExcel, String elemFila, int numRepeticion) para obtener los datos de los trabajadores
                    System.out.println("Añadidos a XML ERRORES: "+ listaDNI.get(i)+" "+ map.get(listaDNI.get(i)));
                     
                    
                    for(int j=2; j< map.get(listaDNI.get(i))+1; j++){
                        
                        try{
                        
                        List<String> filaTrabajador= this.obtenerFilaRepeticiones(localizacionExcel, listaDNI.get(i), j);
                        Trabajador trabajadorProvisional1 = trabajadorAux.rellenarTrabajadorExcel(filaTrabajador);
                        System.out.println(trabajadorProvisional1.getNombre()+" repetido");

                        trabajadoresErrores.add(trabajadorProvisional1);
                        }catch(Exception e){
                            e.printStackTrace();
                        }
                        
                    }
                    
                    listaDNI_Repetidos.add(listaDNI.get(i));
                }
                
                //IGUAL HAY QUE METER ESTO DE DEBAJO EN ALGÚN ELSE, POR LO DE LOS REPETIDOS
                
                int comprobacion=esValidoDNI(listaDNI.get(i));
                
                
                
                
                switch(comprobacion){
                    case 1:
                        //System.out.println("valido: "+listaDNI.get(i));
                        break;
                    case 2:
                        //el error se puede subsanar -> LA LETRA ESTA MAL
                        String dniArreglado = arreglarDNI(listaDNI.get(i));  //DNI CON LA LETRA CORRECTA
                        this.modificarDatos(localizacionExcel, 0, listaDNI.get(i), dniArreglado);
                        System.out.println("El dni: "+listaDNI.get(i)+" ha sido reemplazado por "+dniArreglado);
                        break;
                    case 3:
                        //el error no es subsanable -> ESTÁ MAL ESTRUCTURADO -> añadir al XML
                        
                        List<String> filaTrabajador= this.obtenerFila(localizacionExcel, listaDNI.get(i));

                        System.out.println(listaDNI.get(i)+"movida longitud "+filaTrabajador.size());

                        
                        
                        try{
                        Trabajador trabajadorProvisional = trabajadorAux.rellenarTrabajadorExcel(filaTrabajador);
                        System.out.println(trabajadorProvisional.getNombre());

                        trabajadoresErrores.add(trabajadorProvisional);
                        }catch(Exception e){
                            e.printStackTrace();
                        }
                        
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
        File archivoXML = new File("C:/Users/Torre/Documents/GitHub/Practica_SI/NominasSI/src/resources/errores.xml");
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        
        DocumentBuilder db = dbf.newDocumentBuilder();
       // Document doc = db.parse(archivoXML);
       
        Document doc = db.newDocument();
        Element rootElement = doc.createElement("Trabajadores");
        doc.appendChild(rootElement);

        // obtenemos la raíz del documento existente
        Element eRaiz = doc.getDocumentElement();

        // creamos un nuevo elemento para cada trabajador
        for (int i = 0; i < trabajadores.size(); i++) {

            Element xmlTrabajador = doc.createElement("Trabajador");

            // creamos el atributo "id" para el trabajador
            Attr atributoID = doc.createAttribute("id");
            atributoID.setValue(""+trabajadores.get(i).getIdTrabajador());
            xmlTrabajador.setAttributeNode(atributoID);

            // creamos los elementos secundarios para el trabajador y les asignamos el valor correspondiente
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


    



    //HASTA AQUI LO LEGAL CUIDADO TETE
    
    
    
    
   
    

    
      public void extraerDatos(String campo) throws FileNotFoundException, IOException {
                
                File f = new File(localizacionExcel);

              
                InputStream inp = new FileInputStream(f);      //poner try catch

                
		XSSFWorkbook workBook = new XSSFWorkbook(inp);   //representa libro excel en formato XLSX
                
		XSSFSheet hoja =  workBook.getSheetAt(0);        //escogemos pagina 1 del excel
		
		Iterator iteradorHoja = hoja.rowIterator();                //creamos iterador para fila de una hoja
		List cellTemp = new ArrayList();
                
                int contador = 1, tope = 0, bloqueo = 0, k = 0, l = 0;
                
		while(iteradorHoja.hasNext()) 
		{
			XSSFRow fila = (XSSFRow) iteradorHoja.next();     //representa una fila de una hoja
			Iterator iteratorFila = fila.cellIterator();          //iterador para celdas de una fila de la hoja
				while(iteratorFila.hasNext())
				{
					XSSFCell celda = (XSSFCell) iteratorFila.next();     //representa la celda actual
                                        if(celda.toString().equals(campo) && bloqueo == 0)
                                        {
                                            tope = contador;
                                            bloqueo = 1;
                                        }                                             //si no coincide con el campo indicado se pasa a otra fila
                                        if(bloqueo == 1 && k == 1)
                                        {
                                            if(fila.getCell(tope-1) != null && l == 0)
                                            {
                                                cellTemp.add(fila.getCell(tope-1).toString());
                                                l = 1;
                                            }else if(fila.getCell(tope-1) == null && l == 0)
                                            {
                                                cellTemp.add("");
                                                l = 1;
                                            }
                                        }
                                        contador++;
				}
                        k = 1;
                        l = 0;
		}
            
            for (Object element : cellTemp) {
                System.out.println(element.toString());
                
                 
            }

                
          /*
          if(campo.equals("Nombre"))
          {
                this.nombres = cellTemp;
          }else if(campo.equals("Apellido1"))
          {
                this.apellido1 = cellTemp;
          }else if(campo.equals("Apellido2"))
          {
                this.apellido2 = cellTemp;
          }else if(campo.equals("Nombre empresa"))
          {
                this.empresa = cellTemp;
          }else if(campo.equals("Categoria"))
          {
              this.categoria = cellTemp;
          }else if(campo.equals("Email"))
          {
              this.email = cellTemp;
          
        */
          
          
          
      }
    
    public void leerExcel() throws FileNotFoundException, IOException, ParserConfigurationException, TransformerException {
		File f = new File(localizacionExcel);
                InputStream inp = new FileInputStream(f);
		XSSFWorkbook workBook = new XSSFWorkbook(inp);
		XSSFSheet hs =  workBook.getSheetAt(2);
                
                List nombresErroneos = new ArrayList<String>();
                List apellido1Erroneos = new ArrayList<String>();
                List apellido2Erroneos = new ArrayList<String>();
                List empresaErroneos = new ArrayList<String>();
                List categoriaErroneos = new ArrayList<String>();
		
		Iterator rowIter = hs.rowIterator();
		List cellTemp = new ArrayList();
                
                int contador = 1, tope = 0, bloqueo = 0, k = 0, l = 0;
                
		while(rowIter.hasNext()) 
		{
			XSSFRow hr = (XSSFRow) rowIter.next();
			Iterator iterator = hr.cellIterator();
				while(iterator.hasNext())
				{
					XSSFCell hcel = (XSSFCell) iterator.next();
                                        if(hcel.toString().equals("NIF/NIE") && bloqueo == 0)
                                        {
                                            tope = contador;
                                            bloqueo = 1;
                                        }
                                        if(bloqueo == 1 && k == 1)
                                        {
                                            if(hr.getCell(tope-1) != null && l == 0)
                                            {
                                                cellTemp.add(hr.getCell(tope-1));
                                                l = 1;
                                            }else if(hr.getCell(tope-1) == null && l == 0)
                                            {
                                                cellTemp.add("");
                                                l = 1;
                                            }
                                        }
                                        contador++;
				}
                        k = 1;
                        l = 0;
		}
        
        int a = 0, suma = 0, id = 0, conta = 0;
        ArrayList<Integer> listaId = new ArrayList<Integer>();
         
         for(int i = 0;i < cellTemp.size();i++)
        {
            if(cellTemp.get(i).toString().equals("") == false && cellTemp.get(i).toString().charAt(0) != 'X' && cellTemp.get(i).toString().charAt(0) != 'Y' && cellTemp.get(i).toString().charAt(0) != 'Z')
            {
             for(int j = 0;j < cellTemp.size();j++)
                {
                if(cellTemp.get(i).toString().equals(cellTemp.get(j).toString()))
                {
                    conta = conta + 1;
       
                    if(conta > 1 && i >= j)
                    {
                        listaId.add(i+2);
                        nombresErroneos.add(this.nombres.get(i));
                        apellido1Erroneos.add(this.apellido1.get(i));
                        if(this.apellido2.get(i).equals("") != true)
                        {
                            apellido2Erroneos.add(this.apellido2.get(i));
                        }else
                        {
                            apellido2Erroneos.add("");
                        }
                        empresaErroneos.add(this.empresa.get(i));
                        categoriaErroneos.add(this.categoria.get(i));
                        }
                }
                }
                
                conta = 0;
                
            for(int j = 0;j < 8;j++)
            {
                a = Character.getNumericValue(cellTemp.get(i).toString().charAt(j));
                suma = suma + a;
         
            }
            if(cellTemp.get(i).toString().charAt(8) != this.devuelveLetra(suma))
            {
                String celda = cellTemp.get(i).toString();
                StringBuffer cadena = new StringBuffer();
                 for(int j = 0;j < 8;j++)
                {
                    cadena.append(cellTemp.get(i).toString().charAt(j));
                }
                cadena.append(this.devuelveLetra(suma));
                this.modificaExcel(celda , cadena.toString());
            }
            suma = 0;
            
            }else if(cellTemp.get(i).toString().equals("") == false)
            {
                for(int j = 0;j < cellTemp.size();j++)
                {
                if(cellTemp.get(i).toString().equals(cellTemp.get(j).toString()))
                {
                    conta = conta + 1;
       
                    if(conta > 1 && i >= j)
                    {
                        listaId.add(i+2);
                        nombresErroneos.add(this.nombres.get(i));
                        apellido1Erroneos.add(this.apellido1.get(i));
                        if(this.apellido2.get(i).equals("") != true)
                        {
                            apellido2Erroneos.add(this.apellido2.get(i));
                        }else
                        {
                            apellido2Erroneos.add("");
                        }
                        empresaErroneos.add(this.empresa.get(i));
                        categoriaErroneos.add(this.categoria.get(i));
                        }
                }
                }
                
                conta = 0;
                
                if(cellTemp.get(i).toString().charAt(0) == 'X' && cellTemp.get(i).toString().charAt(8) != '0' )
                {
                    StringBuffer cadena = new StringBuffer();
                    String celda = cellTemp.get(i).toString();
                    for(int j = 0;j < 8;j++)
                    {
                        cadena.append(cellTemp.get(i).toString().charAt(j));
                    }
                    cadena.append("0");
                    this.modificaExcel(celda , cadena.toString());
                }else if (cellTemp.get(i).toString().charAt(0) == 'Y' && cellTemp.get(i).toString().charAt(8) != '1')
                {
                    StringBuffer cadena = new StringBuffer();
                    String celda = cellTemp.get(i).toString();
                    for(int j = 0;j < 8;j++)
                    {
                        cadena.append(cellTemp.get(i).toString().charAt(j));
                    }
                    cadena.append("1");
                    this.modificaExcel(celda , cadena.toString());
                }else if(cellTemp.get(i).toString().charAt(0) == 'Z' && cellTemp.get(i).toString().charAt(8) != '2')
                {
                    StringBuffer cadena = new StringBuffer();
                    String celda = cellTemp.get(i).toString();
                    for(int j = 0;j < 8;j++)
                    {
                        cadena.append(cellTemp.get(i).toString().charAt(j));
                    }
                    cadena.append("2");
                    this.modificaExcel(celda , cadena.toString());
                }
            }else if(cellTemp.get(i).toString().equals("") == true)
            {
                if(this.nombres.get(i).equals("") != true)
                {
                    id = i+2;
                    listaId.add(id);
                    nombresErroneos.add(this.nombres.get(i));
                    apellido1Erroneos.add(this.apellido1.get(i));
                    if(this.apellido2.get(i).equals("") != true)
                    {
                        apellido2Erroneos.add(this.apellido2.get(i));
                    }else
                    {
                        apellido2Erroneos.add("");
                    }
                    empresaErroneos.add(this.empresa.get(i));
                    categoriaErroneos.add(this.categoria.get(i));
                }
            }
        }
         
        try {
            this.rellenarXML(listaId,nombresErroneos,apellido1Erroneos,apellido2Erroneos,empresaErroneos,categoriaErroneos);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ExcelManager.class.getName()).log(Level.SEVERE, null, ex);
        } catch (TransformerConfigurationException ex) {
            Logger.getLogger(ExcelManager.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void rellenarXML(ArrayList<Integer> idList,List<String> nombresErroneos,List<String> apellido1Erroneos,List<String> apellido2Erroneos,List<String> empresasErroneos,List<String> categoriasErroneos) throws ParserConfigurationException, FileNotFoundException, IOException, TransformerConfigurationException, TransformerException{
          // Archivo XML
          
            DocumentBuilderFactory dbf   = DocumentBuilderFactory.newInstance();
            DocumentBuilder        db    = dbf.newDocumentBuilder();
            Document               doc   = db.newDocument();
            Element                eRaiz = doc.createElement("Trabajadores");
            
            doc.appendChild(eRaiz);
            
            for(int i = 0;i < nombresErroneos.size();i++)
            {
            
                        Element xmlCuentaIDvalue = doc.createElement("Trabajador");

                        eRaiz.appendChild(xmlCuentaIDvalue);

                        Attr atributoIDdeCuenta = doc.createAttribute("id");

                        atributoIDdeCuenta.setValue(idList.get(i).toString());
                        xmlCuentaIDvalue.setAttributeNode(atributoIDdeCuenta);

                        Element nombreUsuarioCuenta = doc.createElement("Nombre");
             
                        nombreUsuarioCuenta.appendChild(doc.createTextNode(nombresErroneos.get(i)));
                        xmlCuentaIDvalue.appendChild(nombreUsuarioCuenta);

                        Element apellido1UsuarioCuenta = doc.createElement("PrimerApellido");

                        apellido1UsuarioCuenta.appendChild(doc.createTextNode(apellido1Erroneos.get(i)));
                        xmlCuentaIDvalue.appendChild(apellido1UsuarioCuenta);
                        
                        Element apellido2UsuarioCuenta = doc.createElement("SegundoApellido");

                        apellido2UsuarioCuenta.appendChild(doc.createTextNode(apellido2Erroneos.get(i)));
                        xmlCuentaIDvalue.appendChild(apellido2UsuarioCuenta);

                        Element cuentaUsuarioEmpresaNombre = doc.createElement("Empresa");

                        cuentaUsuarioEmpresaNombre.appendChild(doc.createTextNode(empresasErroneos.get(i)));
                        xmlCuentaIDvalue.appendChild(cuentaUsuarioEmpresaNombre);

                        Element categoria1 = doc.createElement("Categoria");

                        categoria1.appendChild(doc.createTextNode(categoriasErroneos.get(i)));
                        xmlCuentaIDvalue.appendChild(categoria1);
                }
                     
        // en el xml
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer        transformer        = transformerFactory.newTransformer();
        DOMSource          source             = new DOMSource(doc);
        StreamResult       result             = new StreamResult(new File("resources/Errores.xml"));
        
        transformer.transform(source, result);
    }
    
    public void modificaExcel(String celda,String nueva) throws FileNotFoundException, IOException{
                File f = new File(localizacionExcel);
                InputStream inp = new FileInputStream(f);
		XSSFWorkbook workBook = new XSSFWorkbook(inp);
		XSSFSheet hs =  workBook.getSheetAt(2);
		
		Iterator rowIter = hs.rowIterator();
                
                int contador = 1, tope = 0, bloqueo = 0, k = 0, l = 0;
                
		while(rowIter.hasNext()) 
		{
			XSSFRow hr = (XSSFRow) rowIter.next();
			Iterator iterator = hr.cellIterator();
				while(iterator.hasNext())
				{
					XSSFCell hcel = (XSSFCell) iterator.next();
                                        if(hcel.toString().equals("NIF/NIE") && bloqueo == 0)
                                        {
                                            tope = contador;
                                            bloqueo = 1;
                                        }
                                        if(bloqueo == 1 && k == 1)
                                        {
                                            if(hr.getCell(tope-1) != null && l == 0 && hr.getCell(tope-1).toString().equals(celda))
                                            {
                                                CellStyle cellStyle = workBook.createCellStyle();
                                                //System.out.println("entro");
                                                hr.getCell(tope-1).setCellValue(nueva);
                                                l = 1;
                                            }
                                        }
                                        contador++;
				}
                        k = 1;
                        l = 0;
		}
        inp.close();
        //Open FileOutputStream to write updates
        FileOutputStream output_file = new FileOutputStream(new File(localizacionExcel));
        //write changes
        workBook.write(output_file);
        //close the stream
        output_file.close();            
    }
    
     public void modificaExcelEmail() throws FileNotFoundException, IOException{
                File f = new File(localizacionExcel);
                InputStream inp = new FileInputStream(f);
		XSSFWorkbook workBook = new XSSFWorkbook(inp);
		XSSFSheet hs =  workBook.getSheetAt(2);
		
		Iterator rowIter = hs.rowIterator();
                
                int contador = 1, tope = 0, bloqueo = 0, k = 0, l = 0, index = 0;
                
		while(rowIter.hasNext()) 
		{
			XSSFRow hr = (XSSFRow) rowIter.next();
			Iterator iterator = hr.cellIterator();
				while(iterator.hasNext())
				{
					XSSFCell hcel = (XSSFCell) iterator.next();
                                        if(hcel.toString().equals("Email") && bloqueo == 0)
                                        {
                                            tope = contador;
                                            bloqueo = 1;
                                        }
                                        if(bloqueo == 1 && k == 1)
                                        {
                                            if(l == 0)
                                            {
                                                if(hr.getCell(tope-1) == null && index < this.listaEmails.size())
                                                {
                                                    hr.createCell(tope-1);
                                                    hr.getCell(tope-1).setCellValue(this.listaEmails.get(index));    
                                                    index++;
                                                    l = 1;
                                                    break;
                                                }else if(index < this.listaEmails.size())
                                                {
                                                    hr.getCell(tope-1).setCellValue(this.listaEmails.get(index));    
                                                    index++;
                                                    l = 1; 
                                                }
                                            }
                                        }
                                        contador++;
				}
                        k = 1;
                        l = 0;
		}
        inp.close();
        //Open FileOutputStream to write updates
        FileOutputStream output_file = new FileOutputStream(new File(localizacionExcel));
        //write changes
        workBook.write(output_file);
        //close the stream
        output_file.close();            
     }
}