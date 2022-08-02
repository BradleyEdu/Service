package proyecto_serviciosocial;

import static java.awt.Color.black;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

/**
 *
 * @author Eliud Ayala
 */
public class panelCons extends javax.swing.JPanel {

    conexion conect = new conexion();
    Connection con = conect.getConnection();

    public static DefaultTableModel modelo = new DefaultTableModel();
    public static DefaultTableModel modelo2 = new DefaultTableModel();

    SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
    SimpleDateFormat fechaBien = new SimpleDateFormat("yyyy-MM-dd");

    Editar edit = new Editar();
    EditarTab2 edit2 = new EditarTab2();
    Agregar agr = new Agregar();

    public int otra;

    public panelCons() {
        initComponents();

        modelo.addColumn("Motivo");
        modelo.addColumn("Del");
        modelo.addColumn("Al");
        modelo.addColumn("Nombre");
        modelo.addColumn("Codigo");
        modelo.addColumn("Ramo registrado");
        modelo.addColumn("Pagaduria");
        modelo.addColumn("Sueldo");
        modelo.addColumn("Quinquenio");
        modelo.addColumn("Otras percepciones");
        modelo.addColumn("Total");
        tablaPuestos.setModel(modelo);

        modelo2.addColumn("Del");
        modelo2.addColumn("Al");
        modelo2.addColumn("Nombre");
        modelo2.addColumn("Codigo");
        modelo2.addColumn("Ramo registrado");
        modelo2.addColumn("Pagaduria");
        modelo2.addColumn("Sueldo");
        modelo2.addColumn("Quinquenio");
        modelo2.addColumn("Otras percepciones");
        modelo2.addColumn("Total");

        this.setSize(500, 500);
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        txtFili = new javax.swing.JTextField();
        txtCurp = new javax.swing.JTextField();
        txtRfc = new javax.swing.JTextField();
        txtHomo = new javax.swing.JTextField();
        btnBusc = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        tablaPuestos = new javax.swing.JTable();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        btnLimpi = new javax.swing.JButton();
        btnImprimir = new javax.swing.JButton();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jLabel8 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        jSeparator1 = new javax.swing.JSeparator();
        jLabel10 = new javax.swing.JLabel();
        jLabel12 = new javax.swing.JLabel();
        jLabel13 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel18 = new javax.swing.JLabel();
        jSeparator2 = new javax.swing.JSeparator();
        jLabel19 = new javax.swing.JLabel();
        jLabel20 = new javax.swing.JLabel();
        jLabel21 = new javax.swing.JLabel();
        jLabel22 = new javax.swing.JLabel();
        jLabel23 = new javax.swing.JLabel();
        jLabel24 = new javax.swing.JLabel();
        jSeparator3 = new javax.swing.JSeparator();
        jScrollPane2 = new javax.swing.JScrollPane();
        tablaTodo = new javax.swing.JTable();
        jLabel11 = new javax.swing.JLabel();
        btnAgregar1 = new javax.swing.JButton();
        btnEditar1 = new javax.swing.JButton();
        btnEliminar1 = new javax.swing.JButton();
        jLabel14 = new javax.swing.JLabel();
        btnAgregar2 = new javax.swing.JButton();
        btnEditar2 = new javax.swing.JButton();
        btnEliminar2 = new javax.swing.JButton();
        txtNom = new javax.swing.JTextField();
        txtPaterno = new javax.swing.JTextField();
        txtMaterno = new javax.swing.JTextField();
        txtCalle = new javax.swing.JTextField();
        txtNumExt = new javax.swing.JTextField();
        txtCol = new javax.swing.JTextField();
        txtCp = new javax.swing.JTextField();
        txtCiudad = new javax.swing.JTextField();
        txtEstado = new javax.swing.JTextField();
        txtIngreso_dia = new javax.swing.JTextField();
        txtIngreso_mes = new javax.swing.JTextField();
        txtIngreso_ano = new javax.swing.JTextField();
        txtBaja_dia = new javax.swing.JTextField();
        txtBaja_mes = new javax.swing.JTextField();
        txtBaja_ano = new javax.swing.JTextField();
        jLabel15 = new javax.swing.JLabel();
        jLabel16 = new javax.swing.JLabel();
        jLabel17 = new javax.swing.JLabel();
        jLabel25 = new javax.swing.JLabel();
        jLabel26 = new javax.swing.JLabel();
        jLabel27 = new javax.swing.JLabel();
        txtMotivos = new javax.swing.JTextField();
        btnActualizar = new javax.swing.JButton();
        jLabel28 = new javax.swing.JLabel();
        txtIngreso_letra = new javax.swing.JTextField();
        jLabel29 = new javax.swing.JLabel();
        txtBaja_letra = new javax.swing.JTextField();

        setBackground(new java.awt.Color(255, 255, 255));
        setMaximumSize(new java.awt.Dimension(700, 650));
        setPreferredSize(new java.awt.Dimension(680, 500));

        txtFili.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtCurp.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtRfc.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtHomo.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        txtHomo.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtHomoActionPerformed(evt);
            }
        });

        btnBusc.setFont(new java.awt.Font("Arial", 0, 18)); // NOI18N
        btnBusc.setText("Buscar persona");
        btnBusc.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnBuscActionPerformed(evt);
            }
        });

        tablaPuestos.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        tablaPuestos.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {},
                {},
                {},
                {}
            },
            new String [] {

            }
        ));
        tablaPuestos.setPreferredSize(new java.awt.Dimension(305, 750));
        jScrollPane1.setViewportView(tablaPuestos);

        jLabel1.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel1.setText("Datos personales");

        jLabel2.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel2.setText("Nombres");

        jLabel3.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel3.setText("Apellido Paterno");

        btnLimpi.setFont(new java.awt.Font("Arial", 0, 18)); // NOI18N
        btnLimpi.setText("Limpiar todo");
        btnLimpi.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnLimpiActionPerformed(evt);
            }
        });

        btnImprimir.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        btnImprimir.setText("Imprimir documento");
        btnImprimir.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnImprimirActionPerformed(evt);
            }
        });

        jLabel6.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel6.setText("Filiacion");

        jLabel7.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel7.setText("Curp");

        jLabel8.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel8.setText("RFC");

        jLabel9.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel9.setText("Homoclave");

        jSeparator1.setBackground(new java.awt.Color(0, 0, 0));
        jSeparator1.setForeground(new java.awt.Color(0, 0, 0));

        jLabel10.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel10.setText("Fecha de ingreso");

        jLabel12.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel12.setText("Motivos de la baja");

        jLabel13.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel13.setText("Fecha de baja");

        jLabel4.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel4.setText("Apellido Materno");

        jLabel5.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel5.setText("Domicilio");

        jLabel18.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel18.setText("Calle");

        jSeparator2.setBackground(new java.awt.Color(0, 0, 0));
        jSeparator2.setForeground(new java.awt.Color(0, 0, 0));

        jLabel19.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel19.setText("Numero Exterior");

        jLabel20.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel20.setText("Colonia o Fraccionamiento");

        jLabel21.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel21.setText("Código postal");

        jLabel22.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel22.setText("Ciudad");

        jLabel23.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel23.setText("Estado");

        jLabel24.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel24.setText("Periodo de aportaciones al fondo del ISSTE");

        jSeparator3.setBackground(new java.awt.Color(0, 0, 0));
        jSeparator3.setForeground(new java.awt.Color(0, 0, 0));

        tablaTodo.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        tablaTodo.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {},
                {},
                {},
                {}
            },
            new String [] {

            }
        ));
        tablaTodo.setPreferredSize(new java.awt.Dimension(305, 750));
        jScrollPane2.setViewportView(tablaTodo);

        jLabel11.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel11.setText("Motivo y periodo en que ocurrió la(s) reingreso(s), o licencia(s)");

        btnAgregar1.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        btnAgregar1.setText("Agregar");
        btnAgregar1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAgregar1ActionPerformed(evt);
            }
        });

        btnEditar1.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        btnEditar1.setText("Editar");
        btnEditar1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnEditar1ActionPerformed(evt);
            }
        });

        btnEliminar1.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        btnEliminar1.setText("Eliminar fila");
        btnEliminar1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnEliminar1ActionPerformed(evt);
            }
        });

        jLabel14.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel14.setText("Percepciones que aportaron al fondo del ISSTE");

        btnAgregar2.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        btnAgregar2.setText("Agregar");
        btnAgregar2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnAgregar2ActionPerformed(evt);
            }
        });

        btnEditar2.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        btnEditar2.setText("Editar");
        btnEditar2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnEditar2ActionPerformed(evt);
            }
        });

        btnEliminar2.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        btnEliminar2.setText("Eliminar fila");
        btnEliminar2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnEliminar2ActionPerformed(evt);
            }
        });

        txtNom.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtPaterno.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtMaterno.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtCalle.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtNumExt.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtCol.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtCp.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtCiudad.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtEstado.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtIngreso_dia.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtIngreso_mes.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtIngreso_ano.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtBaja_dia.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtBaja_mes.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtBaja_ano.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        jLabel15.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel15.setText("Dia");

        jLabel16.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel16.setText("Mes");

        jLabel17.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel17.setText("Año");

        jLabel25.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel25.setText("Dia");

        jLabel26.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel26.setText("Mes");

        jLabel27.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel27.setText("Año");

        txtMotivos.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        btnActualizar.setBackground(new java.awt.Color(204, 255, 204));
        btnActualizar.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        btnActualizar.setText("<html>\n<body>\n<h3>Actualizar <br>Información</h3>\n</body>\n</html>");
        btnActualizar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnActualizarActionPerformed(evt);
            }
        });

        jLabel28.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel28.setText("Fecha de ingreso con letra");

        txtIngreso_letra.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        txtIngreso_letra.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtIngreso_letraActionPerformed(evt);
            }
        });

        jLabel29.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        jLabel29.setText("Fecha de baja con letra");

        txtBaja_letra.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N
        txtBaja_letra.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                txtBaja_letraActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1)
                    .addComponent(jScrollPane2)
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel11)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jLabel12)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                .addComponent(txtMotivos, javax.swing.GroupLayout.PREFERRED_SIZE, 301, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(btnActualizar, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(150, 150, 150))
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                .addGap(0, 0, Short.MAX_VALUE)
                                .addComponent(jSeparator3, javax.swing.GroupLayout.PREFERRED_SIZE, 742, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel10)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtIngreso_dia, javax.swing.GroupLayout.PREFERRED_SIZE, 41, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(jLabel15))
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtIngreso_mes, javax.swing.GroupLayout.PREFERRED_SIZE, 41, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(jLabel16))
                                        .addGap(15, 15, 15)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addGroup(layout.createSequentialGroup()
                                                .addGap(10, 10, 10)
                                                .addComponent(jLabel17))
                                            .addComponent(txtIngreso_ano, javax.swing.GroupLayout.PREFERRED_SIZE, 66, javax.swing.GroupLayout.PREFERRED_SIZE))))
                                .addGap(49, 49, 49)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel13)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                            .addGroup(layout.createSequentialGroup()
                                                .addComponent(txtBaja_dia, javax.swing.GroupLayout.PREFERRED_SIZE, 41, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                                .addComponent(txtBaja_mes, javax.swing.GroupLayout.PREFERRED_SIZE, 41, javax.swing.GroupLayout.PREFERRED_SIZE))
                                            .addGroup(layout.createSequentialGroup()
                                                .addComponent(jLabel25)
                                                .addGap(28, 28, 28)
                                                .addComponent(jLabel26)))
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtBaja_ano, javax.swing.GroupLayout.PREFERRED_SIZE, 66, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(jLabel27))))
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jLabel28)
                                        .addGap(106, 106, 106))
                                    .addGroup(layout.createSequentialGroup()
                                        .addGap(116, 116, 116)
                                        .addComponent(jLabel29)
                                        .addGap(0, 0, Short.MAX_VALUE))
                                    .addGroup(layout.createSequentialGroup()
                                        .addGap(24, 24, 24)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtBaja_letra)
                                            .addComponent(txtIngreso_letra)))))
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel24)
                                    .addComponent(jLabel8)
                                    .addComponent(txtRfc, javax.swing.GroupLayout.PREFERRED_SIZE, 212, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtFili, javax.swing.GroupLayout.PREFERRED_SIZE, 202, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(jLabel6))
                                        .addGap(78, 78, 78)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel7)
                                            .addGroup(layout.createSequentialGroup()
                                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                            .addComponent(jLabel9)
                                                            .addComponent(txtHomo, javax.swing.GroupLayout.PREFERRED_SIZE, 283, javax.swing.GroupLayout.PREFERRED_SIZE))
                                                        .addGap(102, 102, 102))
                                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                                        .addComponent(txtCurp, javax.swing.GroupLayout.PREFERRED_SIZE, 213, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                        .addGap(172, 172, 172)))
                                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                    .addComponent(btnBusc)
                                                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                                        .addComponent(btnLimpi)
                                                        .addGap(13, 13, 13))))))
                                    .addComponent(jLabel5)
                                    .addComponent(jLabel1)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel2)
                                            .addComponent(txtNom, javax.swing.GroupLayout.PREFERRED_SIZE, 289, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addGap(62, 62, 62)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel3)
                                            .addComponent(txtPaterno, javax.swing.GroupLayout.PREFERRED_SIZE, 194, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addGap(36, 36, 36)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtMaterno, javax.swing.GroupLayout.PREFERRED_SIZE, 194, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(jLabel4)))
                                    .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 794, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 794, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtCalle, javax.swing.GroupLayout.PREFERRED_SIZE, 334, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(txtCp, javax.swing.GroupLayout.PREFERRED_SIZE, 176, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(jLabel21))
                                        .addGap(45, 45, 45)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(txtNumExt, javax.swing.GroupLayout.PREFERRED_SIZE, 176, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addGroup(layout.createSequentialGroup()
                                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                    .addComponent(jLabel22)
                                                    .addComponent(txtCiudad, javax.swing.GroupLayout.PREFERRED_SIZE, 176, javax.swing.GroupLayout.PREFERRED_SIZE))
                                                .addGap(44, 44, 44)
                                                .addComponent(txtEstado, javax.swing.GroupLayout.PREFERRED_SIZE, 176, javax.swing.GroupLayout.PREFERRED_SIZE))))
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(jLabel18)
                                        .addGap(335, 335, 335)
                                        .addComponent(jLabel19)
                                        .addGap(71, 71, 71)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(jLabel20)
                                            .addComponent(txtCol, javax.swing.GroupLayout.PREFERRED_SIZE, 264, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(jLabel23))))
                                .addGap(0, 0, Short.MAX_VALUE))))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(319, 319, 319)
                        .addComponent(btnAgregar1)
                        .addGap(51, 51, 51)
                        .addComponent(btnEditar1)
                        .addGap(47, 47, 47)
                        .addComponent(btnEliminar1))
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jLabel14))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(324, 324, 324)
                        .addComponent(btnAgregar2)
                        .addGap(51, 51, 51)
                        .addComponent(btnEditar2)
                        .addGap(47, 47, 47)
                        .addComponent(btnEliminar2))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(411, 411, 411)
                        .addComponent(btnImprimir)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel6)
                            .addComponent(jLabel7))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(txtFili, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtCurp, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(17, 17, 17)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel8)
                            .addComponent(jLabel9))
                        .addGap(5, 5, 5)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(txtRfc, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(txtHomo, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(41, 41, 41)
                        .addComponent(btnBusc)
                        .addGap(18, 18, 18)
                        .addComponent(btnLimpi)))
                .addGap(34, 34, 34)
                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 17, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 2, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel4)
                            .addComponent(jLabel3))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(txtPaterno)
                            .addComponent(txtMaterno, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jLabel2)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(txtNom, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(18, 18, 18)
                .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 17, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 2, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel19)
                    .addComponent(jLabel18)
                    .addComponent(jLabel20))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtCalle)
                    .addComponent(txtNumExt)
                    .addComponent(txtCol))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel21)
                    .addComponent(jLabel22)
                    .addComponent(jLabel23))
                .addGap(10, 10, 10)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                    .addComponent(txtEstado, javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(txtCiudad, javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(txtCp))
                .addGap(18, 18, Short.MAX_VALUE)
                .addComponent(jLabel24, javax.swing.GroupLayout.PREFERRED_SIZE, 17, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jSeparator3, javax.swing.GroupLayout.PREFERRED_SIZE, 2, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(19, 19, 19)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel10)
                    .addComponent(jLabel13)
                    .addComponent(jLabel28))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtIngreso_dia, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtIngreso_mes, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtIngreso_ano, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtBaja_dia, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtBaja_mes, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtBaja_ano, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(txtIngreso_letra, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel15)
                    .addComponent(jLabel16)
                    .addComponent(jLabel17)
                    .addComponent(jLabel25)
                    .addComponent(jLabel26)
                    .addComponent(jLabel27)
                    .addComponent(jLabel29))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(txtBaja_letra, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(22, 22, 22)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel12)
                            .addComponent(txtMotivos, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(18, 18, 18)
                        .addComponent(jLabel11))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(btnActualizar, javax.swing.GroupLayout.PREFERRED_SIZE, 49, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 148, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnAgregar1)
                    .addComponent(btnEditar1)
                    .addComponent(btnEliminar1))
                .addGap(22, 22, 22)
                .addComponent(jLabel14)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 132, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnAgregar2)
                    .addComponent(btnEditar2)
                    .addComponent(btnEliminar2))
                .addGap(18, 18, 18)
                .addComponent(btnImprimir))
        );

        getAccessibleContext().setAccessibleParent(this);
    }// </editor-fold>//GEN-END:initComponents

    private void btnImprimirActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnImprimirActionPerformed
        try {
            llenarNombre();
        } catch (IOException ex) {
            Logger.getLogger(panelCons.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_btnImprimirActionPerformed

    private void txtHomoActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtHomoActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtHomoActionPerformed

    private void btnBuscActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnBuscActionPerformed
        try {
            consultarPersonal();
            JOptionPane.showMessageDialog(null, "Si hay campos vacíos, completelos y presione el botón ACTUALIZAR INFORMACION", "COMPLETAR LOS CAMPOS", JOptionPane.WARNING_MESSAGE);
        } catch (IOException ex) {
            Logger.getLogger(panelCons.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_btnBuscActionPerformed

    private void btnAgregar1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAgregar1ActionPerformed
        agr.setVisible(true);
    }//GEN-LAST:event_btnAgregar1ActionPerformed

    private void btnEditar1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnEditar1ActionPerformed
        int fila = tablaPuestos.getSelectedRow();

        edit.fila = edit.recibeFila(fila);
        edit.modelo = edit.recibeModelo(modelo);

        edit.lblfila.setText(String.valueOf(fila));
        otra = fila;
        System.out.println("La fila seleccionada es: " + fila);
        System.out.println("Otra vale: " + otra);

        if (fila >= 0) {
            String motivo = (String) modelo.getValueAt(fila, 0);

            String del = (String) modelo.getValueAt(fila, 1);
            String diaD, mesD, anoD;
            diaD = del.substring(0, 2);
            mesD = del.substring(3, 5);
            anoD = del.substring(6, 10);

            String al = (String) modelo.getValueAt(fila, 2);
            String diaA, mesA, anoA;
            diaA = al.substring(0, 2);
            mesA = al.substring(3, 5);
            anoA = al.substring(6, 10);

            String nombre = (String) modelo.getValueAt(fila, 3);
            String ramo = (String) modelo.getValueAt(fila, 4);
            String codigo = (String) modelo.getValueAt(fila, 5);
            String pagad = (String) modelo.getValueAt(fila, 6);
            String sueldo = (String) modelo.getValueAt(fila, 7);
            Double quinquenio = (Double) modelo.getValueAt(fila, 8);
            Double otras = (Double) modelo.getValueAt(fila, 9);
            Double total = (Double) modelo.getValueAt(fila, 10);

            edit.txtMotivo.setText(motivo);
            edit.txtDel_dia.setText(diaD);
            edit.txtDel_mes.setText(mesD);
            edit.txtDel_ano.setText(anoD);
            edit.txtAl_dia.setText(diaA);
            edit.txtAl_mes.setText(mesA);
            edit.txtAl_ano.setText(anoA);
            edit.txtNom.setText(nombre);
            edit.txtCod.setText(codigo);
            edit.txtRamo.setText(ramo);
            edit.txtPagad.setText(pagad);
            edit.txtSueldo.setText(sueldo.toString());
            edit.txtQuinq.setText(quinquenio.toString());
            edit.txtOtras.setText(otras.toString());
            edit.txtTotal.setText(total.toString());

            edit.setVisible(true);
        } else {
            JOptionPane.showMessageDialog(null, "Seleccione una fila");
        }
    }//GEN-LAST:event_btnEditar1ActionPerformed

    private void btnEliminar1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnEliminar1ActionPerformed
        if (tablaPuestos.getSelectedRow() >= 1) {
            int fila = tablaPuestos.getSelectedRow();
            modelo.removeRow(fila);
        } else {
            JOptionPane.showMessageDialog(null, "Seleccione una fila");
        }

    }//GEN-LAST:event_btnEliminar1ActionPerformed

    private void btnEditar2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnEditar2ActionPerformed
        int fila = tablaTodo.getSelectedRow();

        edit2.fila = edit.recibeFila(fila);
        edit2.modelo = edit.recibeModelo(modelo2);

        if (fila >= 0) {
            String del = (String) modelo2.getValueAt(fila, 0);
            String diaD, mesD, anoD;
            diaD = del.substring(0, 2);
            mesD = del.substring(3, 5);
            anoD = del.substring(6, 10);

            String al = (String) modelo2.getValueAt(fila, 1);
            String diaA, mesA, anoA;
            diaA = al.substring(0, 2);
            mesA = al.substring(3, 5);
            anoA = al.substring(6, 10);
            String nombre = (String) modelo2.getValueAt(fila, 2);
            String ramo = (String) modelo2.getValueAt(fila, 3);
            String codigo = (String) modelo2.getValueAt(fila, 4);
            String pagad = (String) modelo2.getValueAt(fila, 5);
            Double sueldo = (Double) modelo2.getValueAt(fila, 6);
            Double quinquenio = (Double) modelo2.getValueAt(fila, 7);
            Double otras = (Double) modelo2.getValueAt(fila, 8);
            Double total = (Double) modelo2.getValueAt(fila, 9);

            edit2.txtDel_dia.setText(diaD);
            edit2.txtDel_mes.setText(mesD);
            edit2.txtDel_ano.setText(anoD);
            edit2.txtAl_dia.setText(diaA);
            edit2.txtAl_mes.setText(mesA);
            edit2.txtAl_ano.setText(anoA);
            edit2.txtNom.setText(nombre);
            edit2.txtRamo.setText(ramo);
            edit2.txtPagad.setText(pagad);
            edit2.txtSueldo.setText(sueldo.toString());
            edit2.txtQuinq.setText(quinquenio.toString());
            edit2.txtOtras.setText(otras.toString());
            edit2.txtTotal.setText(total.toString());

            edit2.setVisible(true);
        } else {
            JOptionPane.showMessageDialog(null, "Seleccione una fila");
        }
    }//GEN-LAST:event_btnEditar2ActionPerformed

    private void btnEliminar2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnEliminar2ActionPerformed
        if (tablaTodo.getSelectedRow() >= 1) {
            int fila = tablaTodo.getSelectedRow();
            modelo2.removeRow(fila);
        } else
            JOptionPane.showMessageDialog(null, "Seleccione una fila");
    }//GEN-LAST:event_btnEliminar2ActionPerformed

    private void btnLimpiActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnLimpiActionPerformed
        txtFili.setText(null);
        txtRfc.setText(null);
        txtCurp.setText(null);
        txtHomo.setText(null);
        txtFili.setText(null);
        txtNom.setText(null);
        txtPaterno.setText(null);
        txtMaterno.setText(null);
        txtCalle.setText(null);
        txtNumExt.setText(null);
        txtCol.setText(null);
        txtCp.setText(null);
        txtCiudad.setText(null);
        txtEstado.setText(null);
        txtIngreso_dia.setText(null);
        txtIngreso_mes.setText(null);
        txtIngreso_ano.setText(null);
        txtBaja_dia.setText(null);
        txtBaja_mes.setText(null);
        txtBaja_ano.setText(null);
        txtMotivos.setText(null);

        int filas1 = tablaPuestos.getRowCount();
        for (int i = filas1 - 1; i >= 0; i--) {
            modelo.removeRow(0);
        }

        int filas2 = tablaTodo.getRowCount();
        for (int i = filas2 - 1; i >= 0; i--) {
            modelo2.removeRow(0);
        }

    }//GEN-LAST:event_btnLimpiActionPerformed

    private void btnAgregar2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnAgregar2ActionPerformed
        Agregar2 agre = new Agregar2();
        agre.setVisible(true);
    }//GEN-LAST:event_btnAgregar2ActionPerformed

    private void btnActualizarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnActualizarActionPerformed

        String filiacion, curp, rfc, homoclave, apellido_paterno, apellido_materno, nombre, calle, numero, colonia,
                codigo_postal, ciudad, estado, ingreso = "", ingreso_letra, baja = "", baja_letra, motivo = "";

        String fechaIng = "", fechaBaja = "";

        filiacion = txtFili.getText();
        curp = txtCurp.getText();
        rfc = txtRfc.getText();
        homoclave = txtHomo.getText();
        apellido_paterno = txtPaterno.getText();
        apellido_materno = txtMaterno.getText();
        nombre = txtNom.getText();
        calle = txtCalle.getText();
        numero = txtNumExt.getText();
        colonia = txtCol.getText();
        codigo_postal = txtCp.getText();
        ciudad = txtCiudad.getText();
        estado = txtEstado.getText();
        fechaIng = txtIngreso_ano.getText() + "-" + txtIngreso_mes.getText() + "-" + txtIngreso_dia.getText();
        ingreso = fechaIng.replace(" ", "");
        ingreso_letra = txtIngreso_letra.getText();
        fechaBaja = txtBaja_ano.getText() + "-" + txtBaja_mes.getText() + "-" + txtBaja_dia.getText();
        baja = fechaBaja.replace(" ", "");
        baja_letra = txtBaja_letra.getText();
        motivo = txtMotivos.getText();

        try {
            System.out.println("---INICIA REGISTRO DE DATOS----");

            String sql = "update personal set filiacion= ?,"
                    + "curp= ?,"
                    + "rfc= ?,"
                    + "homoclave= ?,"
                    + "apellido_paterno= ?,"
                    + "apellido_materno= ?,"
                    + "nombre= ?,"
                    + "calle= ?,"
                    + "numero= ?,"
                    + "colonia= ?,"
                    + "codigo_postal= ?,"
                    + "ciudad= ?,"
                    + "estado= ?,"
                    + "ingreso= ?,"
                    + "ingreso_letra= ?,"
                    + "baja= ?,"
                    + "baja_letra= ?,"
                    + "motivo_baja= ?"
                    + "where filiacion='" + filiacion + "'";

            PreparedStatement ps = con.prepareStatement(sql);
            ps.setString(1, filiacion);
            ps.setString(2, curp);
            ps.setString(3, rfc);
            ps.setString(4, homoclave);
            ps.setString(5, apellido_paterno);
            ps.setString(6, apellido_materno);
            ps.setString(7, nombre);
            ps.setString(8, calle);
            ps.setString(9, numero);
            ps.setString(10, colonia);
            ps.setString(11, codigo_postal);
            ps.setString(12, ciudad);
            ps.setString(13, estado);
            ps.setString(14, ingreso);
            ps.setString(15, ingreso_letra);
            ps.setString(16, baja);
            ps.setString(17, baja_letra);
            ps.setString(18, motivo);
            ps.executeUpdate();

            JOptionPane.showMessageDialog(null, "Registro Actualizado");
            //limpiar();
        } catch (SQLException ex) {
            Logger.getLogger(panelCons.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("---------------No se actualizo-----------------");
        }
    }//GEN-LAST:event_btnActualizarActionPerformed

    private void txtIngreso_letraActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtIngreso_letraActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtIngreso_letraActionPerformed

    private void txtBaja_letraActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_txtBaja_letraActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_txtBaja_letraActionPerformed

    public Date sacarFechaBaja(String filia) {
        try {
            String sql = "SELECT del FROM personal_sueldo WHERE motivo != '" + "" + "' AND "
                    + "filiacion = '" + filia + "' ORDER BY del ASC LIMIT 1";

            Statement sentencia = con.createStatement();
            ResultSet rs = sentencia.executeQuery(sql);

            while (rs.next()) {
                Date fechaBaja = rs.getDate(1);

                if (rs.getDate(1) != null) {
                    return fechaBaja;
                }
            }

        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Error de consulta", "Error en la consulta de la fecha", JOptionPane.ERROR_MESSAGE);
        }
        return null;
    }

    public void consultarPersonal() throws IOException {

        try {
            String sql = "SELECT curp, rfc, homoclave, apellido_paterno, apellido_materno"
                    + ", nombre, calle, numero, colonia, codigo_postal, ciudad, estado"
                    + ", ingreso, ingreso_letra, baja, baja_letra, motivo_baja FROM personal WHERE filiacion = '" + txtFili.getText() + "'";

            Statement sentencia = con.createStatement();
            ResultSet rs = sentencia.executeQuery(sql);

            while (rs.next()) {
                txtCurp.setText(rs.getString(1));
                txtRfc.setText(rs.getString(2));
                txtHomo.setText(rs.getString(3));
                txtPaterno.setText(rs.getString(4));
                txtMaterno.setText(rs.getString(5));
                txtNom.setText(rs.getString(6));
                txtCalle.setText(rs.getString(7));
                txtNumExt.setText(rs.getString(8));
                txtCol.setText(rs.getString(9));
                txtCp.setText(rs.getString(10));
                txtCiudad.setText(rs.getString(11));
                txtEstado.setText(rs.getString(12));

                Date fechaIngreso = rs.getDate(13);
                String ingreso = (formatter.format(fechaIngreso));

                txtIngreso_dia.setText(ingreso.substring(0, 2));
                txtIngreso_mes.setText(ingreso.substring(3, 5));
                txtIngreso_ano.setText(ingreso.substring(6, 10));
                txtIngreso_letra.setText((rs.getString(14)));

                if (rs.getDate(15) != null) {
                    Date fechaBaja = rs.getDate(15);
                    String baja = (formatter.format(fechaBaja));
                    txtBaja_dia.setText(baja.substring(0, 2));
                    txtBaja_mes.setText(baja.substring(3, 5));
                    txtBaja_ano.setText(baja.substring(6, 10));
                }

                txtBaja_letra.setText(rs.getString(16));
                txtMotivos.setText(rs.getString(17));
            }

        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Error de consulta", "Error en la consulta de datos generales", JOptionPane.ERROR_MESSAGE);
        }

        consultarPersonal_Sueldo();
        llenarTabla2();
    }

    public void consultarPersonal_Sueldo() throws IOException {
        Object columnas[] = new Object[12];
        String sql = "";
        try {

            if (sacarFechaBaja(txtFili.getText()) != null) {

                sql = "SELECT motivo, del, al, puesto, codigo, nivel, sueldo, quinquenio, otras, "
                        + "sum(sueldo + quinquenio + otras) as total FROM personal_sueldo "
                        + "WHERE filiacion = '" + txtFili.getText() + "' "
                        + "AND del >= '" + sacarFechaBaja(txtFili.getText()) + "' "
                        //+ "AND motivo != '"+""+"' "
                        + "GROUP BY motivo, del, al, puesto, codigo, nivel, sueldo, quinquenio, otras"
                        + " ORDER BY del ASC";
            } else {

                sql = "SELECT motivo, del, al, puesto, codigo, nivel, sueldo, quinquenio, otras, "
                        + "sum(sueldo + quinquenio + otras) as total FROM personal_sueldo "
                        + "WHERE filiacion = '" + txtFili.getText() + "' "
                        //+ "AND del >= '" + sacarFechaBaja(txtFili.getText()) + "' "
                        + "AND motivo != '" + "" + "' "
                        + "GROUP BY motivo, del, al, puesto, codigo, nivel, sueldo, quinquenio, otras"
                        + " ORDER BY del ASC";
            }

            Statement sentencia = con.createStatement();
            ResultSet rs = sentencia.executeQuery(sql);

            while (rs.next()) {
                columnas[0] = rs.getString(1);

                Date fechaDel = rs.getDate(2);
                Date fechaAl = rs.getDate(3);

                columnas[1] = formatter.format(fechaDel);
                if (fechaAl == null) {
                    columnas[2] = "";
                } else {
                    columnas[2] = formatter.format(fechaAl);
                }
                columnas[3] = rs.getString(4);
                columnas[4] = rs.getString(5);
                columnas[5] = rs.getString(6);
                columnas[6] = "";
                columnas[7] = rs.getString(7);
                columnas[8] = rs.getDouble(7);
                columnas[9] = rs.getDouble(8);
                columnas[10] = rs.getDouble(9);
                columnas[11] = rs.getDouble(10);
                modelo.addRow(columnas);
            }

            tablaPuestos.setModel(modelo);

        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Error de consulta", "Tabla"
                    + " de personal_sueldo", JOptionPane.ERROR_MESSAGE);
        }
    }

    public void llenarTabla2() throws IOException {
        Object columnas[] = new Object[12];

        try {
            String sql = "SELECT del, al, puesto, codigo, nivel, sueldo, quinquenio, otras, "
                    + "sum(sueldo + quinquenio + otras) as total FROM personal_sueldo "
                    + "WHERE filiacion = '" + "ROGS671001D40" + "' "
                    + "AND ISNULL(motivo) "
                    + "GROUP BY del, al, puesto, codigo, nivel, sueldo, quinquenio, otras "
                    + "ORDER BY del ASC";

            Statement sentencia = con.createStatement();
            ResultSet rs = sentencia.executeQuery(sql);

            while (rs.next()) {
                Date fechaDel = rs.getDate(1);
                Date fechaAl = rs.getDate(2);
                columnas[0] = formatter.format(fechaDel);
                if (fechaAl == null) {
                    columnas[1] = "";
                } else {
                    columnas[1] = formatter.format(fechaAl);
                }
                columnas[2] = rs.getString(3);
                columnas[3] = rs.getString(4);
                columnas[4] = "";
                columnas[5] = rs.getString(5);
                columnas[6] = rs.getDouble(6);
                columnas[7] = rs.getDouble(7);
                columnas[8] = rs.getDouble(8);
                columnas[9] = rs.getDouble(9);
                modelo2.addRow(columnas);

            }
            tablaTodo.setModel(modelo2);

        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Error de consulta", "Error en la consulta", JOptionPane.ERROR_MESSAGE);
        }
    }

    public void llenarNombre() throws IOException {
        try {
            FileInputStream file = new FileInputStream(new File("C:\\Users\\dog_a\\Downloads\\servicioo_2.0 (1).xlsx"));

            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFSheet sheet2 = wb.getSheetAt(1);

            XSSFRow fila = sheet.getRow(5);
            XSSFRow fila_domicilio = sheet.getRow(11);
            XSSFRow fila_aportaciones = sheet.getRow(19);

            // Comienza el llenado de la parte de informacion de la persona
            if (fila == null) {
                fila = sheet.createRow(5);
            }

            XSSFCell nombre = fila.createCell(6);
            XSSFCell apellido_pat = fila.createCell(0);
            XSSFCell apellido_mat = fila.createCell(3);
            XSSFCell curp = fila.createCell(11);
            XSSFCell rfc = fila.createCell(9);

            if (nombre == null) {
                nombre = fila.createCell(6);
            }

            if (apellido_pat == null) {
                apellido_pat = fila.createCell(0);
            }

            if (apellido_mat == null) {
                apellido_pat = fila.createCell(3);
            }

            if (curp == null) {
                curp = fila.createCell(11);
            }

            if (rfc == null) {
                rfc = fila.createCell(9);
            }

            //Termina el llenado de la informacion de la persona
            // Comienza el llenado de la parte del domicilio
            if (fila_domicilio == null) {
                fila_domicilio = sheet.createRow(11);
            }

            XSSFCell calle = fila_domicilio.createCell(0);
            XSSFCell numExt = fila_domicilio.createCell(4);
            XSSFCell colonia = fila_domicilio.createCell(5);
            XSSFCell cp = fila_domicilio.createCell(9);
            XSSFCell ciudad = fila_domicilio.createCell(10);
            XSSFCell estado = fila_domicilio.createCell(12);

            if (calle == null) {
                calle = fila_domicilio.createCell(0);
            }

            if (numExt == null) {
                numExt = fila_domicilio.createCell(4);
            }

            if (colonia == null) {
                colonia = fila_domicilio.createCell(5);
            }

            if (cp == null) {
                cp = fila_domicilio.createCell(9);
            }

            if (ciudad == null) {
                ciudad = fila_domicilio.createCell(10);
            }

            if (estado == null) {
                estado = fila_domicilio.createCell(12);
            }

            //Termina llenado del domicilio
            //Comienza el llenado de la fila de aportaciones
            if (fila_aportaciones == null) {
                fila_aportaciones = sheet.createRow(19);
            }

            XSSFCell ingreso = fila_aportaciones.createCell(0);
            XSSFCell ingreso_letra = fila_aportaciones.createCell(2);
            XSSFCell baja = fila_aportaciones.createCell(7);
            XSSFCell baja_letra = fila_aportaciones.createCell(9);

            if (ingreso == null) {
                ingreso = fila_aportaciones.createCell(0);
            }

            if (ingreso_letra == null) {
                ingreso_letra = fila_aportaciones.createCell(4);
            }

            if (baja == null) {
                baja = fila_aportaciones.createCell(5);
            }

            if (baja_letra == null) {
                baja_letra = fila_aportaciones.createCell(9);
            }

            //Termina llenado de la fila de aportaciones
            CellStyle cellStyle = fila.getSheet().getWorkbook().createCellStyle();
            cellStyle.setAlignment(CellStyle.ALIGN_CENTER);

            Font cellFont = wb.createFont();
            cellFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
            cellStyle.setFont(cellFont);

            //Se les da estilo a las celdas del xlsx
            nombre.setCellStyle(cellStyle);
            apellido_pat.setCellStyle(cellStyle);
            apellido_mat.setCellStyle(cellStyle);
            curp.setCellStyle(cellStyle);
            rfc.setCellStyle(cellStyle);

            calle.setCellStyle(cellStyle);
            numExt.setCellStyle(cellStyle);
            colonia.setCellStyle(cellStyle);
            cp.setCellStyle(cellStyle);
            ciudad.setCellStyle(cellStyle);
            estado.setCellStyle(cellStyle);

            ingreso.setCellStyle(cellStyle);
            ingreso_letra.setCellStyle(cellStyle);
            baja.setCellStyle(cellStyle);
            baja_letra.setCellStyle(cellStyle);

            nombre.setCellValue(txtNom.getText());
            apellido_pat.setCellValue(txtPaterno.getText());
            apellido_mat.setCellValue(txtMaterno.getText());
            curp.setCellValue(txtCurp.getText());
            rfc.setCellValue(txtRfc.getText());

            calle.setCellValue(txtCalle.getText());
            numExt.setCellValue(txtNumExt.getText());
            colonia.setCellValue(txtCol.getText());
            cp.setCellValue(txtCp.getText());
            ciudad.setCellValue(txtCiudad.getText());
            estado.setCellValue(txtEstado.getText());

            String fechaIng = txtIngreso_dia.getText() + "/" + txtIngreso_mes.getText() + "/" + txtIngreso_ano.getText();
            String ingresoN = fechaIng.replace(" ", "");
            String fechaBaja = txtBaja_dia.getText() + "/" + txtBaja_mes.getText() + "/" + txtBaja_ano.getText();
            String bajaN = fechaBaja.replace(" ", "");

            ingreso.setCellValue(ingresoN);
            ingreso_letra.setCellValue(txtIngreso_letra.getText());
            baja.setCellValue(bajaN);
            baja_letra.setCellValue(txtBaja_letra.getText());

            llenarTablaHD(file, wb, sheet);
            llenarTablaHD2(file, wb, sheet2);

            file.close();

            FileOutputStream output = new FileOutputStream("C:\\Users\\dog_a\\Downloads\\servicioo_2.0 (1).xlsx");
            wb.write(output);
            output.close();

            File file1 = new File("C:\\Users\\dog_a\\Downloads\\servicioo_2.0 (1).xlsx");
            Desktop.getDesktop().open(file1);

        } catch (FileNotFoundException ex) {
            Logger.getLogger(panelCons.class.getName()).log(Level.SEVERE, null, ex);
            JOptionPane.showMessageDialog(null, "Cierre el programa Excel para generar el formato");
        }
    }


    public void llenarTablaHD(FileInputStream file, XSSFWorkbook wb, XSSFSheet sh) throws IOException {

        int limite = tablaPuestos.getRowCount() - 1;
        int j = 0;
        int cont = 28;
        Object datos[] = new Object[11];

        while (j < limite) {

            XSSFRow fila = sh.getRow(cont);

            if (fila == null) {
                fila = sh.createRow(cont);
            }

            XSSFCell celda1 = fila.createCell(0); //Motivo
            XSSFCell celda2_dia = fila.createCell(1); //del_dia
            XSSFCell celda2_mes = fila.createCell(2); //del_mes
            XSSFCell celda2_ano = fila.createCell(3); //del_ano
            XSSFCell celda3_dia = fila.createCell(4); //al
            XSSFCell celda3_mes = fila.createCell(5); //al
            XSSFCell celda3_ano = fila.createCell(6); //al
            XSSFCell celda4 = fila.createCell(7); //nombre
            XSSFCell celda5 = fila.createCell(8); //codigo
            XSSFCell celda6 = fila.createCell(9); //ramo registrado
            XSSFCell celda7 = fila.createCell(10); //pagaduria registrada
            XSSFCell celda8 = fila.createCell(11); //saldo basico
            XSSFCell celda9 = fila.createCell(12); //quinquenio
            XSSFCell celda10 = fila.createCell(13); //otras percepciones
            XSSFCell celda11 = fila.createCell(14); //total

            if (celda1 == null) {
                celda1 = fila.createCell(0);
            }

            if (celda2_dia == null) {
                celda2_dia = fila.createCell(1);
            }

            if (celda2_mes == null) {
                celda2_mes = fila.createCell(2);
            }

            if (celda2_ano == null) {
                celda2_ano = fila.createCell(3);
            }

            if (celda3_dia == null) {
                celda3_dia = fila.createCell(4);
            }

            if (celda3_mes == null) {
                celda3_mes = fila.createCell(5);
            }

            if (celda3_ano == null) {
                celda3_ano = fila.createCell(6);
            }

            if (celda4 == null) {
                celda4 = fila.createCell(7);
            }

            if (celda5 == null) {
                celda5 = fila.createCell(8);
            }

            if (celda6 == null) {
                celda6 = fila.createCell(9);
            }

            if (celda7 == null) {
                celda7 = fila.createCell(10);
            }

            if (celda8 == null) {
                celda8 = fila.createCell(11);
            }

            if (celda9 == null) {
                celda9 = fila.createCell(12);
            }

            if (celda10 == null) {
                celda10 = fila.createCell(13);
            }

            if (celda11 == null) {
                celda11 = fila.createCell(14);
            }

            for (int i = 0; i < datos.length; i++) {

                datos[i] = tablaPuestos.getValueAt(j, i);

                if (datos[i] == null) {
                    datos[i] = "            ";
                }
            }

            CellStyle cellStyle2 = fila.getSheet().getWorkbook().createCellStyle();
            cellStyle2.setAlignment(CellStyle.ALIGN_CENTER);
            cellStyle2.setBorderBottom(BorderStyle.THIN);
            cellStyle2.setBorderLeft(BorderStyle.THIN);
            cellStyle2.setBorderRight(BorderStyle.THIN);

            celda1.setCellStyle(cellStyle2);
            celda2_dia.setCellStyle(cellStyle2);
            celda2_mes.setCellStyle(cellStyle2);
            celda2_ano.setCellStyle(cellStyle2);
            celda3_dia.setCellStyle(cellStyle2);
            celda3_mes.setCellStyle(cellStyle2);
            celda3_ano.setCellStyle(cellStyle2);
            celda4.setCellStyle(cellStyle2);
            celda5.setCellStyle(cellStyle2);
            celda6.setCellStyle(cellStyle2);
            celda7.setCellStyle(cellStyle2);
            celda8.setCellStyle(cellStyle2);
            celda9.setCellStyle(cellStyle2);
            celda10.setCellStyle(cellStyle2);
            celda11.setCellStyle(cellStyle2);

            celda1.setCellValue(datos[0].toString());
            celda2_dia.setCellValue((datos[1].toString()).substring(0, 2));
            celda2_mes.setCellValue((datos[1].toString()).substring(3, 5));
            celda2_ano.setCellValue(datos[1].toString().substring(6, 10));
            celda3_dia.setCellValue(datos[2].toString().substring(0, 2));
            celda3_mes.setCellValue(datos[2].toString().substring(3, 5));
            celda3_ano.setCellValue(datos[2].toString().substring(6, 10));
            celda4.setCellValue(datos[3].toString());
            celda5.setCellValue(datos[4].toString());
            celda6.setCellValue(datos[5].toString());
            celda7.setCellValue(datos[6].toString());
            int punto = datos[7].toString().indexOf(".");
            celda8.setCellValue(datos[7].toString().substring(0, punto + 3));
            celda9.setCellValue(datos[8].toString());
            celda10.setCellValue(datos[9].toString());
            celda11.setCellValue(datos[10].toString());
            j++;
            cont++;
        }
    }

    public void llenarTablaHD2(FileInputStream file, XSSFWorkbook wb, XSSFSheet sh) throws IOException {

        int limite = tablaTodo.getRowCount() - 1;
        int j = 0;
        int cont = 6;
        Object datos[] = new Object[10];

        while (j < limite) {

            XSSFRow fila = sh.getRow(cont);

            if (fila == null) {
                fila = sh.createRow(cont);
            }

            XSSFCell celda2_dia = fila.createCell(0); //del_dia
            XSSFCell celda2_mes = fila.createCell(1); //del_mes
            XSSFCell celda2_ano = fila.createCell(2); //del_ano
            XSSFCell celda3_dia = fila.createCell(3); //al
            XSSFCell celda3_mes = fila.createCell(4); //al
            XSSFCell celda3_ano = fila.createCell(5); //al
            XSSFCell celda4 = fila.createCell(6); //nombre
            XSSFCell celda5 = fila.createCell(7); //codigo
            XSSFCell celda6 = fila.createCell(8); //ramo registrado
            XSSFCell celda7 = fila.createCell(9); //pagaduria registrada
            XSSFCell celda8 = fila.createCell(10); //saldo basico
            XSSFCell celda9 = fila.createCell(11); //quinquenio
            XSSFCell celda10 = fila.createCell(12); //otras percepciones
            XSSFCell celda11 = fila.createCell(13); //total

            if (celda2_dia == null) {
                celda2_dia = fila.createCell(0);
            }

            if (celda2_mes == null) {
                celda2_mes = fila.createCell(1);
            }

            if (celda2_ano == null) {
                celda2_ano = fila.createCell(2);
            }

            if (celda3_dia == null) {
                celda3_dia = fila.createCell(3);
            }

            if (celda3_mes == null) {
                celda3_mes = fila.createCell(4);
            }

            if (celda3_ano == null) {
                celda3_ano = fila.createCell(5);
            }

            if (celda4 == null) {
                celda4 = fila.createCell(6);
            }

            if (celda5 == null) {
                celda5 = fila.createCell(7);
            }

            if (celda6 == null) {
                celda6 = fila.createCell(8);
            }

            if (celda7 == null) {
                celda7 = fila.createCell(9);
            }

            if (celda8 == null) {
                celda8 = fila.createCell(10);
            }

            if (celda9 == null) {
                celda9 = fila.createCell(11);
            }

            if (celda10 == null) {
                celda10 = fila.createCell(12);
            }

            if (celda11 == null) {
                celda11 = fila.createCell(13);
            }

            for (int i = 0; i < datos.length; i++) {

                datos[i] = tablaTodo.getValueAt(j, i);

                if (datos[i] == null) {
                    datos[i] = "            ";
                }
            }

            celda2_dia.setCellValue((datos[0].toString()).substring(0, 2));
            celda2_mes.setCellValue((datos[0].toString()).substring(3, 5));
            celda2_ano.setCellValue(datos[0].toString().substring(6, 10));
            celda3_dia.setCellValue(datos[1].toString().substring(0, 2));
            celda3_mes.setCellValue(datos[1].toString().substring(3, 5));
            celda3_ano.setCellValue(datos[1].toString().substring(6, 10));
            celda4.setCellValue(datos[2].toString());
            celda5.setCellValue(datos[3].toString());
            celda6.setCellValue(datos[4].toString());
            celda7.setCellValue(datos[5].toString());
            int punto = datos[6].toString().indexOf(".");
            celda8.setCellValue(datos[6].toString().substring(0, punto + 2));
            celda9.setCellValue(datos[7].toString());
            celda10.setCellValue(datos[8].toString());
            celda11.setCellValue(datos[9].toString());
            j++;
            cont++;
        }
        
        insertarImg(file, wb, sh, cont + 10);
    }

    public void insertarImg(FileInputStream file, XSSFWorkbook wb, XSSFSheet sh, int filaF) throws IOException {
                    
    
            XSSFRow fila = sh.getRow(filaF);

            if (fila == null) {
                fila = sh.createRow(filaF);
            }

            XSSFCell celda = fila.createCell(0);

            if (celda == null) {
                celda = fila.createCell(0);
            }

            //FileInputStream obtains input bytes from the image file
            InputStream inputStream = new FileInputStream("C:\\condusef.png");
            
            //Get the contents of an InputStream as a byte[].
            byte[] bytes = IOUtils.toByteArray(inputStream);
            
            //Adds a picture to the workbook
            int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
            
            //close the input stream
            inputStream.close();

            //Returns an object that handles instantiating concrete classes
            CreationHelper helper = wb.getCreationHelper();

            //Creates the top-level drawing patriarch.
            Drawing drawing = sh.createDrawingPatriarch();

            //Create an anchor that is attached to the worksheet
            ClientAnchor anchor = helper.createClientAnchor();
            
            //set top-left corner for the image
            anchor.setCol1(0);
            anchor.setRow1(filaF);

            //Creates a picture
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            //Reset the image to the original size
            pict.resize();
   
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnActualizar;
    private javax.swing.JButton btnAgregar1;
    private javax.swing.JButton btnAgregar2;
    private javax.swing.JButton btnBusc;
    private javax.swing.JButton btnEditar1;
    private javax.swing.JButton btnEditar2;
    private javax.swing.JButton btnEliminar1;
    private javax.swing.JButton btnEliminar2;
    private javax.swing.JButton btnImprimir;
    private javax.swing.JButton btnLimpi;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel14;
    private javax.swing.JLabel jLabel15;
    private javax.swing.JLabel jLabel16;
    private javax.swing.JLabel jLabel17;
    private javax.swing.JLabel jLabel18;
    private javax.swing.JLabel jLabel19;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel20;
    private javax.swing.JLabel jLabel21;
    private javax.swing.JLabel jLabel22;
    private javax.swing.JLabel jLabel23;
    private javax.swing.JLabel jLabel24;
    private javax.swing.JLabel jLabel25;
    private javax.swing.JLabel jLabel26;
    private javax.swing.JLabel jLabel27;
    private javax.swing.JLabel jLabel28;
    private javax.swing.JLabel jLabel29;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JSeparator jSeparator2;
    private javax.swing.JSeparator jSeparator3;
    public static javax.swing.JTable tablaPuestos;
    public static javax.swing.JTable tablaTodo;
    public javax.swing.JTextField txtBaja_ano;
    public javax.swing.JTextField txtBaja_dia;
    public javax.swing.JTextField txtBaja_letra;
    public javax.swing.JTextField txtBaja_mes;
    public javax.swing.JTextField txtCalle;
    public javax.swing.JTextField txtCiudad;
    public javax.swing.JTextField txtCol;
    public javax.swing.JTextField txtCp;
    public javax.swing.JTextField txtCurp;
    public javax.swing.JTextField txtEstado;
    public static javax.swing.JTextField txtFili;
    public javax.swing.JTextField txtHomo;
    public javax.swing.JTextField txtIngreso_ano;
    public javax.swing.JTextField txtIngreso_dia;
    public javax.swing.JTextField txtIngreso_letra;
    public javax.swing.JTextField txtIngreso_mes;
    public javax.swing.JTextField txtMaterno;
    public javax.swing.JTextField txtMotivos;
    public javax.swing.JTextField txtNom;
    public javax.swing.JTextField txtNumExt;
    public javax.swing.JTextField txtPaterno;
    public javax.swing.JTextField txtRfc;
    // End of variables declaration//GEN-END:variables
}
