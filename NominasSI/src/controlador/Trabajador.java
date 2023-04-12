package controlador;
// Generated 14-mar-2023 11:36:10 by Hibernate Tools 4.3.1


import java.util.Date;
import java.util.List;

/**
 * Trabajador generated by hbm2java
 */
public class Trabajador  implements java.io.Serializable {


     private int idTrabajador;
     private String nombre;
     private String apellido1;
     private String apellido2;
     private String nifnie;
     private String email;
     private Date fechaAlta;
     private String codigoCuenta;
     private String iban;
     private Date bajaLaboral;
     private Date altaLaboral;
     private int empresasIdEmpresa;
     private int categoriasIdCategoria;
     private String empresa;
     private String categoria;

    public Trabajador() {
    }
    
    //CONSTRUCTOR PARA MAPEAR TRABAJADORES DEL EXCEL
    
    public Trabajador (int idTrabajador, String codigoCuenta, String iban, String email, Date fechaAlta, String empresa, String categoria, String apellido1, String apellido2, String nombre, String nifnie, Date bajaLaboral, Date altaLaboral) {
        this.idTrabajador = idTrabajador;
        this.codigoCuenta = codigoCuenta;
        this.iban = iban;
        this.email = email;
        this.fechaAlta = fechaAlta;
        this.empresa = empresa;
        this.categoria = categoria;
        this.apellido1 = apellido1;
        this.apellido2 = apellido2;
        this.nombre = nombre;
        this.nifnie = nifnie;
        this.bajaLaboral = bajaLaboral;
        this.altaLaboral = altaLaboral; 
    }
    
    //

	
    public Trabajador(int idTrabajador, String nombre, String apellido1, String nifnie, int empresasIdEmpresa, int categoriasIdCategoria, String empresa, String categoria) {
        this.idTrabajador = idTrabajador;
        this.nombre = nombre;
        this.apellido1 = apellido1;
        this.nifnie = nifnie;
        this.empresasIdEmpresa = empresasIdEmpresa;
        this.categoriasIdCategoria = categoriasIdCategoria;
        this.empresa = empresa;
        this.categoria=categoria;
                
        
        
    }
    public Trabajador(int idTrabajador, String nombre, String apellido1, String apellido2, String nifnie, String email, Date fechaAlta, String codigoCuenta, String iban, Date bajaLaboral, Date altaLaboral, int empresasIdEmpresa, int categoriasIdCategoria) {
       this.idTrabajador = idTrabajador;
       this.nombre = nombre;
       this.apellido1 = apellido1;
       this.apellido2 = apellido2;
       this.nifnie = nifnie;
       this.email = email;
       this.fechaAlta = fechaAlta;
       this.codigoCuenta = codigoCuenta;
       this.iban = iban;
       this.bajaLaboral = bajaLaboral;
       this.altaLaboral = altaLaboral;
       this.empresasIdEmpresa = empresasIdEmpresa;
       this.categoriasIdCategoria = categoriasIdCategoria;
    }
    
    public Trabajador rellenarTrabajadorExcel(List<String> lista){

        // SE ELIMINAN LOS ESPACIOS EN BLANCO DE LA LISTA

        for (int i=0; i<lista.size(); i++) {
            
            if (lista.get(i) == "") {
                lista.remove(i);
                i--;
            }
        }
        
        Trabajador nuevo = new Trabajador();
        
        if(lista.size()==12){
            nuevo.setIdTrabajador(Integer.parseInt(lista.get(11)));
        }else if(lista.size()==11){
            nuevo.setIdTrabajador(Integer.parseInt(lista.get(10)));
        }else if(lista.size()==10) {
            nuevo.setIdTrabajador(Integer.parseInt(lista.get(9)));
        }
        
        nuevo.setNifnie(lista.get(9));
        //System.out.println("EL NIF  ES "+nuevo.getNifnie());
        nuevo.setNombre(lista.get(8));
        //System.out.println("EL nombre en cambio  ES "+nuevo.getNombre());
        nuevo.setApellido1(lista.get(6));
        nuevo.setApellido2(lista.get(7));
        nuevo.setEmpresa(lista.get(4));
        nuevo.setCategoria(lista.get(5));
            
        
        return nuevo;
    }
    
    public String getEmpresa() {
        return this.empresa;
    }
    public String getCategoria() {
        return this.categoria;
    }
    
    public void setEmpresa(String empresa) {
        this.empresa = empresa;
    }
    public void setCategoria(String categoria) {
        this.categoria = categoria;
    }
    public int getIdTrabajador() {
        return this.idTrabajador;
    }
    
    public void setIdTrabajador(int idTrabajador) {
        this.idTrabajador = idTrabajador;
    }
    public String getNombre() {
        return this.nombre;
    }
    
    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
    public String getApellido1() {
        return this.apellido1;
    }
    
    public void setApellido1(String apellido1) {
        this.apellido1 = apellido1;
    }
    public String getApellido2() {
        return this.apellido2;
    }
    
    public void setApellido2(String apellido2) {
        this.apellido2 = apellido2;
    }
    public String getNifnie() {
        return this.nifnie;
    }
    
    public void setNifnie(String nifnie) {
        this.nifnie = nifnie;
    }
    public String getEmail() {
        return this.email;
    }
    
    public void setEmail(String email) {
        this.email = email;
    }
    public Date getFechaAlta() {
        return this.fechaAlta;
    }
    
    public void setFechaAlta(Date fechaAlta) {
        this.fechaAlta = fechaAlta;
    }
    public String getCodigoCuenta() {
        return this.codigoCuenta;
    }
    
    public void setCodigoCuenta(String codigoCuenta) {
        this.codigoCuenta = codigoCuenta;
    }
    public String getIban() {
        return this.iban;
    }
    
    public void setIban(String iban) {
        this.iban = iban;
    }
    public Date getBajaLaboral() {
        return this.bajaLaboral;
    }
    
    public void setBajaLaboral(Date bajaLaboral) {
        this.bajaLaboral = bajaLaboral;
    }
    public Date getAltaLaboral() {
        return this.altaLaboral;
    }
    
    public void setAltaLaboral(Date altaLaboral) {
        this.altaLaboral = altaLaboral;
    }
    public int getEmpresasIdEmpresa() {
        return this.empresasIdEmpresa;
    }
    
    public void setEmpresasIdEmpresa(int empresasIdEmpresa) {
        this.empresasIdEmpresa = empresasIdEmpresa;
    }
    public int getCategoriasIdCategoria() {
        return this.categoriasIdCategoria;
    }
    
    public void setCategoriasIdCategoria(int categoriasIdCategoria) {
        this.categoriasIdCategoria = categoriasIdCategoria;
    }




}


