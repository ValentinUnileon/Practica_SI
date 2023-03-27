package controlador;
// Generated 14-mar-2023 11:36:10 by Hibernate Tools 4.3.1


import java.util.Date;

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

    public Trabajador() {
    }

	
    public Trabajador(int idTrabajador, String nombre, String apellido1, String nifnie, int empresasIdEmpresa, int categoriasIdCategoria) {
        this.idTrabajador = idTrabajador;
        this.nombre = nombre;
        this.apellido1 = apellido1;
        this.nifnie = nifnie;
        this.empresasIdEmpresa = empresasIdEmpresa;
        this.categoriasIdCategoria = categoriasIdCategoria;
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

