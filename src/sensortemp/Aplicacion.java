/*
 * PROYECTO ARDUINO - SENSORES - EXCEL
 */
package sensortemp;

import Arduino.Arduino;
import gnu.io.SerialPortEvent;
import gnu.io.SerialPortEventListener;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.PrintWriter;
import java.util.Calendar;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author HACKTOTOPO
 */
public class Aplicacion extends javax.swing.JFrame {

    Arduino Arduino = new Arduino();
    int Slot=1;
    String temperatura1 = "";
    String temperatura2 = "";
    String temperatura3 = "";
    DefaultTableModel Modelo;
    int Lecturas =0;
                 
    SerialPortEventListener evento = new SerialPortEventListener() {

        @Override
        public void serialEvent(SerialPortEvent spe) {
            if(Arduino.MessageAvailable()== true ){
                
                if(Slot==1){
                    Slot=2;
                    
                    if (Lecturas > 1) {
                        ActualizarTabla();
                    }
                    
                    
                    Lecturas++;
                    temperatura1 = Arduino.PrintMessage();
                }
                
                else if(Slot==2){
                    Slot=3;
                    Lecturas++;
                    temperatura2 = Arduino.PrintMessage();
                }
                
                else if(Slot==3){
                    Slot=1;
                    Lecturas++;
                    temperatura3 = Arduino.PrintMessage();
                
                }
                
            }
        }
    };
    
    boolean State =false;
    Calendar Calendario = Calendar.getInstance();
    
    
    public void ActualizarTabla(){
            String Output = "";
            String hora = Calendario.get(Calendar.HOUR_OF_DAY) + "";
            String minuto = Calendario.get(Calendar.MINUTE) + "";
            String segundos = Calendario.get(Calendar.SECOND)+ "";
            
            if(Integer.parseInt(hora)<10){
                hora = "0" + hora;
            }
            if(Integer.parseInt(minuto)<10){
                minuto = "0" + minuto;
            }
            if(Integer.parseInt(segundos)<10){
                segundos = "0" + segundos;
            }
        
            Output = hora + ":" + minuto + ":" + segundos;
            Calendario = Calendar.getInstance();
            
        Modelo.addRow(new Object[]{Output,temperatura1,temperatura2,temperatura3});
    }
    
    
    public Aplicacion() {
        initComponents();
        this.setLocationRelativeTo(null);
        
        Modelo = (DefaultTableModel) jTable1.getModel();
        try {
            Arduino.ArduinoRXTX("COM4", 2000, 9600, evento);
        } catch (Exception ex) {
            Logger.getLogger(Aplicacion.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void ExcelDatos(String input) {
        HSSFWorkbook libro = new HSSFWorkbook();
        HSSFSheet hoja = libro.createSheet();
        HSSFRow fila = hoja.createRow(0);
        HSSFCell celda = fila.createCell(0);
        celda.setCellValue("VALORES DE TEMPERATURA T1, T2, T3 :");
        fila = hoja.createRow(1);
        celda = fila.createCell(0);
        celda.setCellValue("HORA");
        celda = fila.createCell(1);
        celda.setCellValue("TEMPERATURA 1");
        celda = fila.createCell(2);
        celda.setCellValue("TEMPERATURA 2");
        celda = fila.createCell(3);
        celda.setCellValue("TEMPERATURA 3");
        

        for (int i = 0; i <= Modelo.getRowCount() - 1; i++) {
            fila = hoja.createRow(i + 3);
            for (int j = 0; j <= 3; j++) {
                celda = fila.createCell(j);
                celda.setCellValue(jTable1.getValueAt(i, j).toString());
            }
        }
        try {
            FileOutputStream Fichero = new FileOutputStream(input);
            libro.write(Fichero);
            Fichero.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenu2 = new javax.swing.JMenu();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jButton3 = new javax.swing.JButton();

        jMenu1.setText("File");
        jMenuBar1.add(jMenu1);

        jMenu2.setText("Edit");
        jMenuBar1.add(jMenu2);

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jTable1.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "TIEMPO", "TEMPERATURA 1", "TEMPERATURA 2", "TEMPERATURA 3"
            }
        ));
        jScrollPane1.setViewportView(jTable1);

        jButton1.setText("CARGAR DATOS");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jButton2.setText("EXPORTAR DATOS");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jButton3.setText("LIMPIAR DATOS");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 491, Short.MAX_VALUE)
            .addGroup(layout.createSequentialGroup()
                .addGap(34, 34, 34)
                .addComponent(jButton1)
                .addGap(32, 32, 32)
                .addComponent(jButton3)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jButton2)
                .addGap(44, 44, 44))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 291, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.CENTER)
                    .addComponent(jButton2)
                    .addComponent(jButton3)
                    .addComponent(jButton1))
                .addGap(5, 5, 5))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    
    
    
    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        if (State == true) {
            jButton1.setText("INICIAR CAPTURA");
            State = false;
            try {
                Arduino.SendData("0");
            } catch (Exception ex) {
                Logger.getLogger(Aplicacion.class.getName()).log(Level.SEVERE, null, ex);
            }
        } else {
            State = true;
            jButton1.setText("DETENER CAPTURA");
            try {
                Arduino.SendData("1");
            } catch (Exception ex) {
                Logger.getLogger(Aplicacion.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
      
        javax.swing.JFileChooser Ventana = new javax.swing.JFileChooser();
        String ruta = "";
        try {
            if (Ventana.showSaveDialog(null) == Ventana.APPROVE_OPTION) {
                ruta = Ventana.getSelectedFile().getAbsolutePath() + ".xls";
                ExcelDatos(ruta);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
          
        
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        Modelo.setRowCount(0);
    }//GEN-LAST:event_jButton3ActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Aplicacion.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Aplicacion.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Aplicacion.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Aplicacion.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Aplicacion().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    // End of variables declaration//GEN-END:variables
}
