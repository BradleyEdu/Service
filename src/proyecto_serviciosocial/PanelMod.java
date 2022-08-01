package proyecto_serviciosocial;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
//import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
//import java.sql.Statement;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
//import javax.swing.table.DefaultTableModel;


public class PanelMod extends javax.swing.JPanel {
    
    //Conectar
    conexion conect = new conexion();
    Connection con = conect.getConnection();
    
    public PanelMod() {
        initComponents();
    }


    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        panelRegistro = new javax.swing.JPanel();
        lbFiliacion = new javax.swing.JLabel();
        txtFiliacion = new javax.swing.JTextField();
        lbCurp = new javax.swing.JLabel();
        txtCurp = new javax.swing.JTextField();
        lbRfc = new javax.swing.JLabel();
        txtRfc = new javax.swing.JTextField();
        lbHomo = new javax.swing.JLabel();
        txtHomo = new javax.swing.JTextField();
        lbDatos = new javax.swing.JLabel();
        lbPaterno = new javax.swing.JLabel();
        txtPaterno = new javax.swing.JTextField();
        lbMaterno = new javax.swing.JLabel();
        txtMaterno = new javax.swing.JTextField();
        lbNombre = new javax.swing.JLabel();
        txtNombre = new javax.swing.JTextField();
        lbDomicilio = new javax.swing.JLabel();
        lbCalle = new javax.swing.JLabel();
        txtCalle = new javax.swing.JTextField();
        lbColonia = new javax.swing.JLabel();
        txtColonia = new javax.swing.JTextField();
        lbNumeroExt = new javax.swing.JLabel();
        txtNumeroExt = new javax.swing.JTextField();
        lbCP = new javax.swing.JLabel();
        txtCP = new javax.swing.JTextField();
        lbCiudad = new javax.swing.JLabel();
        txtCiudad = new javax.swing.JTextField();
        lbEstado = new javax.swing.JLabel();
        txtEstado = new javax.swing.JTextField();
        lbNombre5 = new javax.swing.JLabel();
        lbIngreso = new javax.swing.JLabel();
        lbIngresoL = new javax.swing.JLabel();
        txtIngresoL = new javax.swing.JTextField();
        lbBaja = new javax.swing.JLabel();
        lbBajaL = new javax.swing.JLabel();
        txtBajaL = new javax.swing.JTextField();
        lbMotivo = new javax.swing.JLabel();
        btnGuardar = new javax.swing.JButton();
        btnCancelar = new javax.swing.JButton();
        txtDiaIng = new javax.swing.JTextField();
        txtMesIng = new javax.swing.JTextField();
        txtAnoIng = new javax.swing.JTextField();
        lbIngreso1 = new javax.swing.JLabel();
        lbIngreso2 = new javax.swing.JLabel();
        lbIngreso3 = new javax.swing.JLabel();
        lbIngreso4 = new javax.swing.JLabel();
        lbIngreso5 = new javax.swing.JLabel();
        lbIngreso6 = new javax.swing.JLabel();
        txtDiaBaja = new javax.swing.JTextField();
        txtMesBaja = new javax.swing.JTextField();
        txtAnoBaja = new javax.swing.JTextField();
        cbBaja = new javax.swing.JComboBox<>();
        lbMotivo1 = new javax.swing.JLabel();
        lbMotivo2 = new javax.swing.JLabel();
        btnCon = new javax.swing.JButton();

        panelRegistro.setBackground(new java.awt.Color(255, 255, 255));

        lbFiliacion.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbFiliacion.setText("FILIACIÓN:");

        txtFiliacion.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbCurp.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbCurp.setText("CURP:");

        txtCurp.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbRfc.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbRfc.setText("RFC:");

        txtRfc.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbHomo.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbHomo.setText("HOMOCLAVE:");

        txtHomo.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbDatos.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbDatos.setText("Datos Personales:");

        lbPaterno.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbPaterno.setText("Apellido Paterno:");

        txtPaterno.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbMaterno.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbMaterno.setText("Apellido Materno:");

        txtMaterno.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbNombre.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbNombre.setText("Nombre (s):");

        txtNombre.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbDomicilio.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbDomicilio.setText("Domicilio:");

        lbCalle.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbCalle.setText("Calle:");

        txtCalle.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbColonia.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbColonia.setText("Colonia:");

        txtColonia.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbNumeroExt.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbNumeroExt.setText("Numero Ext. e Int:");

        txtNumeroExt.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbCP.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbCP.setText("CP:");

        txtCP.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbCiudad.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbCiudad.setText("Ciudad:");

        txtCiudad.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbEstado.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbEstado.setText("Estado:");

        txtEstado.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbNombre5.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbNombre5.setText("PERIODO DE APORTACIONES AL FONDO DEL ISSSTE:");

        lbIngreso.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso.setText("Ingreso:");

        lbIngresoL.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngresoL.setText("Ingreso con letra:");

        txtIngresoL.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbBaja.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbBaja.setText("Baja:");

        lbBajaL.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbBajaL.setText("Baja con letra:");

        txtBajaL.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbMotivo.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbMotivo.setText("Motivo de baja:");

        btnGuardar.setBackground(new java.awt.Color(255, 255, 255));
        btnGuardar.setForeground(new java.awt.Color(255, 255, 255));
        btnGuardar.setIcon(new javax.swing.ImageIcon(getClass().getResource("/imagenes/guardar.png"))); // NOI18N
        btnGuardar.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 1, 1, 1));
        btnGuardar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnGuardarActionPerformed(evt);
            }
        });

        btnCancelar.setBackground(new java.awt.Color(255, 255, 255));
        btnCancelar.setForeground(new java.awt.Color(255, 255, 255));
        btnCancelar.setIcon(new javax.swing.ImageIcon(getClass().getResource("/imagenes/cancel.png"))); // NOI18N
        btnCancelar.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 1, 1, 1));
        btnCancelar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnCancelarActionPerformed(evt);
            }
        });

        txtDiaIng.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtMesIng.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtAnoIng.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        lbIngreso1.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso1.setText("dd");

        lbIngreso2.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso2.setText("mm");

        lbIngreso3.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso3.setText("aaaa");

        lbIngreso4.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso4.setText("dd");

        lbIngreso5.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso5.setText("mm");

        lbIngreso6.setFont(new java.awt.Font("Arial", 1, 18)); // NOI18N
        lbIngreso6.setText("aaaa");

        txtDiaBaja.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtMesBaja.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        txtAnoBaja.setFont(new java.awt.Font("Arial", 0, 16)); // NOI18N

        cbBaja.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "ABANDONO DE EMPLEO",
            "ACTIVO(A)","DEFUNCIÓN",
            "INCAPACIDAD TOTAL Y PERMANENTE", "JUBILADO(A)",
            "PENSIONADO(A)", "RENUNCIA", "RETIRO VOLUNTARIO" }));

lbMotivo1.setFont(new java.awt.Font("Arial", 1, 16)); // NOI18N
lbMotivo1.setText("Guardar");

lbMotivo2.setFont(new java.awt.Font("Arial", 1, 16)); // NOI18N
lbMotivo2.setText("Cancelar");

btnCon.setIcon(new javax.swing.ImageIcon(getClass().getResource("/imagenes/lu2.png"))); // NOI18N
btnCon.addActionListener(new java.awt.event.ActionListener() {
    public void actionPerformed(java.awt.event.ActionEvent evt) {
        btnConActionPerformed(evt);
    }
    });

    javax.swing.GroupLayout panelRegistroLayout = new javax.swing.GroupLayout(panelRegistro);
    panelRegistro.setLayout(panelRegistroLayout);
    panelRegistroLayout.setHorizontalGroup(
        panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
        .addGroup(panelRegistroLayout.createSequentialGroup()
            .addGap(23, 23, 23)
            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createSequentialGroup()
                    .addGap(0, 0, Short.MAX_VALUE)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createSequentialGroup()
                            .addComponent(btnGuardar, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addGap(56, 56, 56))
                        .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createSequentialGroup()
                            .addComponent(lbMotivo1)
                            .addGap(68, 68, 68)))
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(panelRegistroLayout.createSequentialGroup()
                            .addGap(74, 74, 74)
                            .addComponent(btnCancelar, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGroup(panelRegistroLayout.createSequentialGroup()
                            .addGap(88, 88, 88)
                            .addComponent(lbMotivo2)))
                    .addGap(338, 338, 338))
                .addGroup(panelRegistroLayout.createSequentialGroup()
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                        .addGroup(javax.swing.GroupLayout.Alignment.LEADING, panelRegistroLayout.createSequentialGroup()
                            .addComponent(lbIngreso)
                            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addGroup(panelRegistroLayout.createSequentialGroup()
                                    .addGap(45, 45, 45)
                                    .addComponent(lbIngreso1)
                                    .addGap(33, 33, 33)
                                    .addComponent(lbIngreso2)
                                    .addGap(18, 18, 18)
                                    .addComponent(lbIngreso3))
                                .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createSequentialGroup()
                                    .addGap(32, 32, 32)
                                    .addComponent(txtDiaIng, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(txtMesIng, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(txtAnoIng, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE)))
                            .addGap(78, 78, 78)
                            .addComponent(lbIngresoL)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                            .addComponent(txtIngresoL))
                        .addGroup(panelRegistroLayout.createSequentialGroup()
                            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addComponent(lbDatos)
                                .addGroup(panelRegistroLayout.createSequentialGroup()
                                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(lbPaterno, javax.swing.GroupLayout.Alignment.TRAILING)
                                            .addComponent(lbMaterno))
                                        .addComponent(lbNombre))
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                        .addComponent(txtPaterno, javax.swing.GroupLayout.DEFAULT_SIZE, 211, Short.MAX_VALUE)
                                        .addComponent(txtMaterno)
                                        .addComponent(txtNombre))))
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 69, Short.MAX_VALUE)
                            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                .addGroup(panelRegistroLayout.createSequentialGroup()
                                    .addComponent(lbCiudad)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(txtCiudad, javax.swing.GroupLayout.PREFERRED_SIZE, 134, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGap(18, 18, 18)
                                    .addComponent(lbEstado)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(txtEstado, javax.swing.GroupLayout.PREFERRED_SIZE, 125, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGroup(panelRegistroLayout.createSequentialGroup()
                                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(lbColonia)
                                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                                .addComponent(lbDomicilio)
                                                .addComponent(lbCalle)))
                                        .addComponent(lbNumeroExt))
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(btnCon, javax.swing.GroupLayout.PREFERRED_SIZE, 84, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                            .addComponent(txtCalle)
                                            .addComponent(txtColonia)
                                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, panelRegistroLayout.createSequentialGroup()
                                                .addComponent(txtNumeroExt, javax.swing.GroupLayout.PREFERRED_SIZE, 134, javax.swing.GroupLayout.PREFERRED_SIZE)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 32, Short.MAX_VALUE)
                                                .addComponent(lbCP)
                                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                                .addComponent(txtCP, javax.swing.GroupLayout.PREFERRED_SIZE, 125, javax.swing.GroupLayout.PREFERRED_SIZE)))))))
                        .addGroup(panelRegistroLayout.createSequentialGroup()
                            .addComponent(lbBaja)
                            .addGap(59, 59, 59)
                            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addGroup(panelRegistroLayout.createSequentialGroup()
                                    .addComponent(txtDiaBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(txtMesBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(txtAnoBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addGroup(panelRegistroLayout.createSequentialGroup()
                                    .addGap(13, 13, 13)
                                    .addComponent(lbIngreso4)
                                    .addGap(33, 33, 33)
                                    .addComponent(lbIngreso5)
                                    .addGap(18, 18, 18)
                                    .addComponent(lbIngreso6)))
                            .addGap(108, 108, 108)
                            .addComponent(lbBajaL)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                            .addComponent(txtBajaL)))
                    .addGap(41, 41, 41))
                .addGroup(panelRegistroLayout.createSequentialGroup()
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                            .addGroup(javax.swing.GroupLayout.Alignment.LEADING, panelRegistroLayout.createSequentialGroup()
                                .addComponent(lbFiliacion)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtFiliacion, javax.swing.GroupLayout.PREFERRED_SIZE, 164, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(lbCurp)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtCurp, javax.swing.GroupLayout.PREFERRED_SIZE, 164, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(javax.swing.GroupLayout.Alignment.LEADING, panelRegistroLayout.createSequentialGroup()
                                .addComponent(lbRfc)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtRfc, javax.swing.GroupLayout.PREFERRED_SIZE, 164, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(lbHomo)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(txtHomo, javax.swing.GroupLayout.PREFERRED_SIZE, 164, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addComponent(lbNombre5)
                        .addGroup(panelRegistroLayout.createSequentialGroup()
                            .addComponent(lbMotivo)
                            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                            .addComponent(cbBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 187, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
    );
    panelRegistroLayout.setVerticalGroup(
        panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
        .addGroup(panelRegistroLayout.createSequentialGroup()
            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(panelRegistroLayout.createSequentialGroup()
                    .addGap(23, 23, 23)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbFiliacion)
                        .addComponent(txtFiliacion, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(lbCurp)
                        .addComponent(txtCurp, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGap(21, 21, 21)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbRfc)
                        .addComponent(txtRfc, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(lbHomo)
                        .addComponent(txtHomo, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGroup(panelRegistroLayout.createSequentialGroup()
                    .addGap(42, 42, 42)
                    .addComponent(btnCon, javax.swing.GroupLayout.PREFERRED_SIZE, 70, javax.swing.GroupLayout.PREFERRED_SIZE)))
            .addGap(41, 41, 41)
            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                .addComponent(lbDatos)
                .addComponent(lbDomicilio))
            .addGap(18, 18, 18)
            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(panelRegistroLayout.createSequentialGroup()
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbPaterno)
                        .addComponent(txtPaterno, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGap(18, 18, 18)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbMaterno)
                        .addComponent(txtMaterno, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGap(18, 18, 18)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbNombre)
                        .addComponent(txtNombre, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGroup(panelRegistroLayout.createSequentialGroup()
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbCalle)
                        .addComponent(txtCalle, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGap(18, 18, 18)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbColonia)
                        .addComponent(txtColonia, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGap(18, 18, 18)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbNumeroExt)
                        .addComponent(txtNumeroExt, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(lbCP)
                        .addComponent(txtCP, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))))
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                .addComponent(lbEstado)
                .addComponent(txtEstado, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addComponent(lbCiudad)
                .addComponent(txtCiudad, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addGap(47, 47, 47)
            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                .addGroup(panelRegistroLayout.createSequentialGroup()
                    .addComponent(lbNombre5)
                    .addGap(46, 46, 46)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbIngreso)
                        .addComponent(lbIngresoL)
                        .addComponent(txtIngresoL, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGroup(panelRegistroLayout.createSequentialGroup()
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbIngreso1)
                        .addComponent(lbIngreso2)
                        .addComponent(lbIngreso3))
                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(txtDiaIng, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(txtMesIng, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(txtAnoIng, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))))
            .addGap(18, 18, 18)
            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lbBaja)
                    .addComponent(lbBajaL)
                    .addComponent(txtBajaL, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGroup(panelRegistroLayout.createSequentialGroup()
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(lbIngreso4)
                        .addComponent(lbIngreso5)
                        .addComponent(lbIngreso6))
                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                    .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(txtDiaBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(txtMesBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(txtAnoBaja, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))))
            .addGap(48, 48, 48)
            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                .addComponent(lbMotivo)
                .addComponent(cbBaja, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addGap(55, 55, 55)
            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addComponent(btnCancelar, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addComponent(btnGuardar, javax.swing.GroupLayout.PREFERRED_SIZE, 90, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addGroup(panelRegistroLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addComponent(lbMotivo1)
                .addComponent(lbMotivo2))
            .addContainerGap(19, Short.MAX_VALUE))
    );

    javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
    this.setLayout(layout);
    layout.setHorizontalGroup(
        layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
        .addComponent(panelRegistro, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
    );
    layout.setVerticalGroup(
        layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
        .addComponent(panelRegistro, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
    );
    }// </editor-fold>//GEN-END:initComponents

    private void btnGuardarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnGuardarActionPerformed
        String sql, filiacion,curp,rfc,homoclave,apellido_paterno, apellido_materno,nombre,calle,numero,colonia,
                codigo_postal,ciudad,estado,ingreso="",ingreso_letra,baja="",baja_letra,motivo_baja="";
        String fechaIng="",fechaBaja="";
        filiacion = txtFiliacion.getText();
        curp = txtCurp.getText();
        rfc = txtRfc.getText();
        homoclave = txtHomo.getText();
        apellido_paterno = txtPaterno.getText();
        apellido_materno = txtMaterno.getText();
        nombre = txtNombre.getText();
        calle = txtCalle.getText();
        numero = txtNumeroExt.getText();
        colonia = txtColonia.getText();
        codigo_postal = txtCP.getText();
        ciudad = txtCiudad.getText();
        estado = txtEstado.getText();
        fechaIng=txtAnoIng.getText()+"-"+txtMesIng.getText()+"-"+txtDiaIng.getText();
        ingreso = fechaIng.replace(" ","");
        ingreso_letra = txtIngresoL.getText();
        fechaBaja=txtAnoBaja.getText()+"-"+txtMesBaja.getText()+"-"+txtDiaBaja.getText();
        baja = fechaBaja.replace(" ","");
        baja_letra = txtBajaL.getText();
        
        int opcBaja;
        opcBaja=cbBaja.getSelectedIndex();
        
        if(opcBaja==0){
            motivo_baja="ABANDONO DE EMPLEO";
        }
        if(opcBaja==1){
            motivo_baja=("ACTIVO(A)");
        }
        if(opcBaja==2){
            motivo_baja=("DEFUNCIÓN");
        }
        if(opcBaja==3){
            motivo_baja=("INCAPACIDAD TOTAL Y PERMANENTE");
        }
        if(opcBaja==4){
           motivo_baja=("JUBILADO(A)");
        }
        if(opcBaja==5){
           motivo_baja= ("PENSIONADO(A)");
        }
        if(opcBaja==6){
           motivo_baja= ("RENUNCIA");
        }
        if(opcBaja==7){
            motivo_baja=("RETIRO VOLUNTARIO");
        }
        
        //insertar en la tabla PERSONAL (campos de la tabla)valores (variables)
        
        try {
            System.out.println("---INICIA REGISTRO DE DATOS----");
//            sql = "update personal set (filiacion,curp,rfc,homoclave,apellido_paterno,apellido_materno,nombre,calle,"
//                    + "numero,colonia,codigo_postal,ciudad,estado,ingreso,ingreso_letra,baja,baja_letra,motivo_baja)"
//                    + " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) where filiacion='"+filiacion+"'";

            sql="update personal set filiacion= ?,"
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
                    + "where filiacion='"+filiacion+"'";
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
            ps.setString(18, motivo_baja);
            ps.executeUpdate();

            JOptionPane.showMessageDialog(null, "Registro Actualizado");
            //limpiar();
        } catch (SQLException ex) {
            Logger.getLogger(PanelMod.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("No se actualizo");
        }
        

    }//GEN-LAST:event_btnGuardarActionPerformed

    private void btnCancelarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnCancelarActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_btnCancelarActionPerformed

    private void btnConActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnConActionPerformed
        // TODO add your handling code here:
        String sql, filiacion="",curp="",rfc="",homoclave="",apellido_paterno="", apellido_materno="",
                nombre="",calle="",numero="",colonia="",
                codigo_postal="",ciudad="",estado="",ingreso="",ingreso_letra="",baja="",baja_letra="",motivo_baja="";
        String fechaIng,fechaBaja;
        int banF=0,banC=0,banR=0,banH=0;
       
        
        filiacion = txtFiliacion.getText();
        curp = txtCurp.getText();
        rfc = txtRfc.getText();
        homoclave = txtHomo.getText();
        
        
        if(filiacion.equals("") && curp.equals("") && rfc.equals("") && homoclave.equals("")){
            JOptionPane.showMessageDialog(null,"Por favor agregue informacion en algun de estos campos (Filiación, CURP, RFC, Homoclave"
                    + ") para realizar la consulta");

            System.out.println("estamos vacios");
        }
        if(!filiacion.equals("")){
            banF=1;
              try {
            System.out.println("---------------Aqui inicia--------------");
            Statement st = con.createStatement();
            sql = "select * from personal where filiacion = '"+filiacion+"'";
            ResultSet pst = st.executeQuery(sql);
            while(pst.next()){
                filiacion =pst.getString(2);
                curp =pst.getString(3);
                rfc =pst.getString(4);
                homoclave =pst.getString(5);
                apellido_paterno =pst.getString(6);
                apellido_materno =pst.getString(7);
                nombre =pst.getString(8);
                calle =pst.getString(9);
                numero =pst.getString(10);
                colonia =pst.getString(11);
                codigo_postal =pst.getString(12);
                ciudad =pst.getString(13);
                estado =pst.getString(14);
                ingreso =pst.getString(15);
                ingreso_letra =pst.getString(16);
                baja =pst.getString(17);
                baja_letra =pst.getString(18);
                motivo_baja =pst.getString(19);
            }     
        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("No se pudo hacer la consulta");
        }
        
        txtFiliacion.setText(filiacion);
        txtCurp.setText(curp);
        txtRfc.setText(rfc);
        txtHomo.setText(homoclave);
        txtPaterno.setText(apellido_paterno);
        txtMaterno.setText(apellido_materno);
        txtNombre.setText(nombre);
        txtCalle.setText(calle);
        txtNumeroExt.setText(numero);
        txtColonia.setText(colonia);
        txtCP.setText(codigo_postal);
        txtCiudad.setText(ciudad);
        txtEstado.setText(estado); 
        txtDiaIng.setText(ingreso.substring(8,11));
        txtMesIng.setText(ingreso.substring(5,7));
        txtAnoIng.setText(ingreso.substring(0,4));
        txtIngresoL.setText(ingreso_letra);
        txtDiaBaja.setText(baja.substring(8,11));
        txtMesBaja.setText(baja.substring(5,7));
        txtAnoBaja.setText(baja.substring(0,4));
        txtBajaL.setText(baja_letra);
        if(motivo_baja.equals("ABANDONO DE EMPLEO")){
            cbBaja.setSelectedIndex(0);
        }
        if(motivo_baja.equals("ACTIVO(A)")){
            cbBaja.setSelectedIndex(1);
        }
        if(motivo_baja.equals("DEFUNCIÓN")){
            cbBaja.setSelectedIndex(2);
        }
        if(motivo_baja.equals("INCAPACIDAD TOTAL Y PERMANENTE")){
            cbBaja.setSelectedIndex(3);
        }
        if(motivo_baja.equals("JUBILADO(A)")){
            cbBaja.setSelectedIndex(4);
        }
        if(motivo_baja.equals("PENSIONADO(A)")){
            cbBaja.setSelectedIndex(5);
        }
        if(motivo_baja.equals("RENUNCIA")){
            cbBaja.setSelectedIndex(6);
        }
        if(motivo_baja.equals("RETIRO VOLUNTARIO")){
            cbBaja.setSelectedIndex(7);
        }
           banF=0;   
        }
        
        if(!curp.equals("")){
            banC=1;
            if(curp.equals("S/N" ) && (banF==0 || banR==0) || banH==0){
//           JOptionPane.showMessageDialog(null,"Por favor ingrese otro dato en alguno de estos campos (Filiación, RFC, Homoclave"
//                    + ") para realizar la consulta");
            }if(!curp.equals("S/N")){
            System.out.println(" curp no esta vacio");
                          try {
            System.out.println("---------------Aqui inicia--------------");
            Statement st = con.createStatement();
            sql = "select * from personal where curp = '"+curp+"'";
            ResultSet pst = st.executeQuery(sql);
            while(pst.next()){
                filiacion =pst.getString(2);
                curp =pst.getString(3);
                rfc =pst.getString(4);
                homoclave =pst.getString(5);
                apellido_paterno =pst.getString(6);
                apellido_materno =pst.getString(7);
                nombre =pst.getString(8);
                calle =pst.getString(9);
                numero =pst.getString(10);
                colonia =pst.getString(11);
                codigo_postal =pst.getString(12);
                ciudad =pst.getString(13);
                estado =pst.getString(14);
                ingreso =pst.getString(15);
                ingreso_letra =pst.getString(16);
                baja =pst.getString(17);
                baja_letra =pst.getString(18);
                motivo_baja =pst.getString(19);
            }     
        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("No se pudo hacer la consulta");
        }
        
        txtFiliacion.setText(filiacion);
        txtCurp.setText(curp);
        txtRfc.setText(rfc);
        txtHomo.setText(homoclave);
        txtPaterno.setText(apellido_paterno);
        txtMaterno.setText(apellido_materno);
        txtNombre.setText(nombre);
        txtCalle.setText(calle);
        txtNumeroExt.setText(numero);
        txtColonia.setText(colonia);
        txtCP.setText(codigo_postal);
        txtCiudad.setText(ciudad);
        txtEstado.setText(estado); 
        txtDiaIng.setText(ingreso.substring(8,11));
        txtMesIng.setText(ingreso.substring(5,7));
        txtAnoIng.setText(ingreso.substring(0,4));
        txtIngresoL.setText(ingreso_letra);
        txtDiaBaja.setText(baja.substring(8,11));
        txtMesBaja.setText(baja.substring(5,7));
        txtAnoBaja.setText(baja.substring(0,4));
        txtBajaL.setText(baja_letra);
        if(motivo_baja.equals("ABANDONO DE EMPLEO")){
            cbBaja.setSelectedIndex(0);
        }
        if(motivo_baja.equals("ACTIVO(A)")){
            cbBaja.setSelectedIndex(1);
        }
        if(motivo_baja.equals("DEFUNCIÓN")){
            cbBaja.setSelectedIndex(2);
        }
        if(motivo_baja.equals("INCAPACIDAD TOTAL Y PERMANENTE")){
            cbBaja.setSelectedIndex(3);
        }
        if(motivo_baja.equals("JUBILADO(A)")){
            cbBaja.setSelectedIndex(4);
        }
        if(motivo_baja.equals("PENSIONADO(A)")){
            cbBaja.setSelectedIndex(5);
        }
        if(motivo_baja.equals("RENUNCIA")){
            cbBaja.setSelectedIndex(6);
        }
        if(motivo_baja.equals("RETIRO VOLUNTARIO")){
            cbBaja.setSelectedIndex(7);
        }
            }
            banC=0; 
        }
        
        if(!rfc.equals("")){
            banR=1;
            System.out.println("rfc no esta vacio");
             if(rfc.equals("S/N") && (banF==0 || banC==0 || banH==0)){
//                JOptionPane.showMessageDialog(null,"Por favor ingrese otro dato en alguno de estos campos (Filiación, CURP, Homoclave"
//                    + ") para realizar la consulta");
            }if(!rfc.equals("S/N")){
            System.out.println(" curp no esta vacio");
                          try {
            System.out.println("---------------Aqui inicia--------------");
            Statement st = con.createStatement();
            sql = "select * from personal where rfc = '"+rfc+"'";
            ResultSet pst = st.executeQuery(sql);
            while(pst.next()){
                filiacion =pst.getString(2);
                curp =pst.getString(3);
                rfc =pst.getString(4);
                homoclave =pst.getString(5);
                apellido_paterno =pst.getString(6);
                apellido_materno =pst.getString(7);
                nombre =pst.getString(8);
                calle =pst.getString(9);
                numero =pst.getString(10);
                colonia =pst.getString(11);
                codigo_postal =pst.getString(12);
                ciudad =pst.getString(13);
                estado =pst.getString(14);
                ingreso =pst.getString(15);
                ingreso_letra =pst.getString(16);
                baja =pst.getString(17);
                baja_letra =pst.getString(18);
                motivo_baja =pst.getString(19);
            }     
        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("No se pudo hacer la consulta");
        }
        
        txtFiliacion.setText(filiacion);
        txtCurp.setText(curp);
        txtRfc.setText(rfc);
        txtHomo.setText(homoclave);
        txtPaterno.setText(apellido_paterno);
        txtMaterno.setText(apellido_materno);
        txtNombre.setText(nombre);
        txtCalle.setText(calle);
        txtNumeroExt.setText(numero);
        txtColonia.setText(colonia);
        txtCP.setText(codigo_postal);
        txtCiudad.setText(ciudad);
        txtEstado.setText(estado); 
        txtDiaIng.setText(ingreso.substring(8,11));
        txtMesIng.setText(ingreso.substring(5,7));
        txtAnoIng.setText(ingreso.substring(0,4));
        txtIngresoL.setText(ingreso_letra);
        txtDiaBaja.setText(baja.substring(8,11));
        txtMesBaja.setText(baja.substring(5,7));
        txtAnoBaja.setText(baja.substring(0,4));
        txtBajaL.setText(baja_letra);
        if(motivo_baja.equals("ABANDONO DE EMPLEO")){
            cbBaja.setSelectedIndex(0);
        }
        if(motivo_baja.equals("ACTIVO(A)")){
            cbBaja.setSelectedIndex(1);
        }
        if(motivo_baja.equals("DEFUNCIÓN")){
            cbBaja.setSelectedIndex(2);
        }
        if(motivo_baja.equals("INCAPACIDAD TOTAL Y PERMANENTE")){
            cbBaja.setSelectedIndex(3);
        }
        if(motivo_baja.equals("JUBILADO(A)")){
            cbBaja.setSelectedIndex(4);
        }
        if(motivo_baja.equals("PENSIONADO(A)")){
            cbBaja.setSelectedIndex(5);
        }
        if(motivo_baja.equals("RENUNCIA")){
            cbBaja.setSelectedIndex(6);
        }
        if(motivo_baja.equals("RETIRO VOLUNTARIO")){
            cbBaja.setSelectedIndex(7);
        }
            }
            banR=0;
        }
        
        if(!homoclave.equals("")){
            System.out.println("homo no estan vacios");
            if(homoclave.equals("S/N") && (banF==0 || banR==0 || banC==0)){
//                JOptionPane.showMessageDialog(null,"Por favor ingrese otro dato en alguno de estos campos (Filiación, RFC, Homoclave"
//                    + ") para realizar la consulta");
            }if(!homoclave.equals("S/N") ){
            System.out.println(" curp no esta vacio");
                          try {
            System.out.println("---------------Aqui inicia--------------");
            Statement st = con.createStatement();
            sql = "select * from personal where homoclave = '"+homoclave+"'";
            ResultSet pst = st.executeQuery(sql);
            while(pst.next()){
                filiacion =pst.getString(2);
                curp =pst.getString(3);
                rfc =pst.getString(4);
                homoclave =pst.getString(5);
                apellido_paterno =pst.getString(6);
                apellido_materno =pst.getString(7);
                nombre =pst.getString(8);
                calle =pst.getString(9);
                numero =pst.getString(10);
                colonia =pst.getString(11);
                codigo_postal =pst.getString(12);
                ciudad =pst.getString(13);
                estado =pst.getString(14);
                ingreso =pst.getString(15);
                ingreso_letra =pst.getString(16);
                baja =pst.getString(17);
                baja_letra =pst.getString(18);
                motivo_baja =pst.getString(19);
            }     
        } catch (SQLException ex) {
            Logger.getLogger(prueba.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("No se pudo hacer la consulta");
        }
        
        txtFiliacion.setText(filiacion);
        txtCurp.setText(curp);
        txtRfc.setText(rfc);
        txtHomo.setText(homoclave);
        txtPaterno.setText(apellido_paterno);
        txtMaterno.setText(apellido_materno);
        txtNombre.setText(nombre);
        txtCalle.setText(calle);
        txtNumeroExt.setText(numero);
        txtColonia.setText(colonia);
        txtCP.setText(codigo_postal);
        txtCiudad.setText(ciudad);
        txtEstado.setText(estado); 
        txtDiaIng.setText(ingreso.substring(8,11));
        txtMesIng.setText(ingreso.substring(5,7));
        txtAnoIng.setText(ingreso.substring(0,4));
        txtIngresoL.setText(ingreso_letra);
        txtDiaBaja.setText(baja.substring(8,11));
        txtMesBaja.setText(baja.substring(5,7));
        txtAnoBaja.setText(baja.substring(0,4));
        txtBajaL.setText(baja_letra);
        if(motivo_baja.equals("ABANDONO DE EMPLEO")){
            cbBaja.setSelectedIndex(0);
        }
        if(motivo_baja.equals("ACTIVO(A)")){
            cbBaja.setSelectedIndex(1);
        }
        if(motivo_baja.equals("DEFUNCIÓN")){
            cbBaja.setSelectedIndex(2);
        }
        if(motivo_baja.equals("INCAPACIDAD TOTAL Y PERMANENTE")){
            cbBaja.setSelectedIndex(3);
        }
        if(motivo_baja.equals("JUBILADO(A)")){
            cbBaja.setSelectedIndex(4);
        }
        if(motivo_baja.equals("PENSIONADO(A)")){
            cbBaja.setSelectedIndex(5);
        }
        if(motivo_baja.equals("RENUNCIA")){
            cbBaja.setSelectedIndex(6);
        }
        if(motivo_baja.equals("RETIRO VOLUNTARIO")){
            cbBaja.setSelectedIndex(7);
        }
            }
            banH=0;
        }
        
    }//GEN-LAST:event_btnConActionPerformed


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnCancelar;
    private javax.swing.JButton btnCon;
    private javax.swing.JButton btnGuardar;
    private javax.swing.JComboBox<String> cbBaja;
    private javax.swing.JLabel lbBaja;
    private javax.swing.JLabel lbBajaL;
    private javax.swing.JLabel lbCP;
    private javax.swing.JLabel lbCalle;
    private javax.swing.JLabel lbCiudad;
    private javax.swing.JLabel lbColonia;
    private javax.swing.JLabel lbCurp;
    private javax.swing.JLabel lbDatos;
    private javax.swing.JLabel lbDomicilio;
    private javax.swing.JLabel lbEstado;
    private javax.swing.JLabel lbFiliacion;
    private javax.swing.JLabel lbHomo;
    private javax.swing.JLabel lbIngreso;
    private javax.swing.JLabel lbIngreso1;
    private javax.swing.JLabel lbIngreso2;
    private javax.swing.JLabel lbIngreso3;
    private javax.swing.JLabel lbIngreso4;
    private javax.swing.JLabel lbIngreso5;
    private javax.swing.JLabel lbIngreso6;
    private javax.swing.JLabel lbIngresoL;
    private javax.swing.JLabel lbMaterno;
    private javax.swing.JLabel lbMotivo;
    private javax.swing.JLabel lbMotivo1;
    private javax.swing.JLabel lbMotivo2;
    private javax.swing.JLabel lbNombre;
    private javax.swing.JLabel lbNombre5;
    private javax.swing.JLabel lbNumeroExt;
    private javax.swing.JLabel lbPaterno;
    private javax.swing.JLabel lbRfc;
    private javax.swing.JPanel panelRegistro;
    private javax.swing.JTextField txtAnoBaja;
    private javax.swing.JTextField txtAnoIng;
    private javax.swing.JTextField txtBajaL;
    private javax.swing.JTextField txtCP;
    private javax.swing.JTextField txtCalle;
    private javax.swing.JTextField txtCiudad;
    private javax.swing.JTextField txtColonia;
    private javax.swing.JTextField txtCurp;
    private javax.swing.JTextField txtDiaBaja;
    private javax.swing.JTextField txtDiaIng;
    private javax.swing.JTextField txtEstado;
    private javax.swing.JTextField txtFiliacion;
    private javax.swing.JTextField txtHomo;
    private javax.swing.JTextField txtIngresoL;
    private javax.swing.JTextField txtMaterno;
    private javax.swing.JTextField txtMesBaja;
    private javax.swing.JTextField txtMesIng;
    private javax.swing.JTextField txtNombre;
    private javax.swing.JTextField txtNumeroExt;
    private javax.swing.JTextField txtPaterno;
    private javax.swing.JTextField txtRfc;
    // End of variables declaration//GEN-END:variables
}
