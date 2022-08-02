package proyecto_serviciosocial;

import java.sql.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class prueba {
    
    public static void main(String [] args){
    conexion conect = new conexion();
    Connection con = conect.getConnection();
        try {
            System.out.println("---------------Aqui inicia--------------");
            Statement st = con.createStatement();
            String sql = "insert into personal (filiacion) values (?)";
            PreparedStatement pst = con.prepareStatement(sql);
            pst.setString(1,"ABCDEFGHI");
            pst.executeUpdate();
            System.out.println("Se agrego a la tabla");
   
        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("NO se guardo");
        }
   
}
   

}
    