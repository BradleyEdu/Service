package proyecto_serviciosocial;

import java.sql.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;

public class conexion {
    Connection ccn = null;
    Statement st = null;
    
    public conexion(){
        String rutafile = "C:\\Users\\dog_a\\Documents\\Servicio_Social\\BD_Servicio\\ISSSTE 2014_be.accdb";
        String Url = "jdbc:ucanaccess://" + rutafile;
        
        try {
            ccn= DriverManager.getConnection(Url);
            st = ccn.createStatement();
            System.out.println("Se hizo la conexion");
        } catch (SQLException ex) {
            Logger.getLogger(conexion.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null,"Conexion Erronea" + ex);
        }
        
    }
    public Connection getConnection(){
        return ccn;
    }
}
