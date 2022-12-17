package com.example.excel;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.Settings;
import android.view.View;
import android.widget.EditText;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

public class MainActivity extends AppCompatActivity {

    //tt

    private EditText edt_input_site_code;
    private File filePath = new File(Environment.getExternalStorageDirectory() + "/"+ edt_input_site_code + "Table_one.xls");

    String [] strings = new String[]{"site_loc_code", "tx_link_type", "tx_dt_no", "tx_eqpt", "tx_eqpt_name", "Tx_Eqpt_WiFi_Name", "hotel_bbu_loc",
            "hotel_bbu_ne_no", "hotel_wdm_loc", "sector_no", "band", "rru_name", "rru_type", "rru_alias", "rru_pair_mimo",
            "hsn", "hpn", "no_of_link",	"dcu_name", "rru_pwr_supp",	"fiber_mode",	"fiber_length",	"fiber_count",
            "bts_name",	"bts_rack",
            "cpri",	"ant_id1",	"ant_id2",	"ant_id3",	"ant_id4",	"ant_id5",	"pole1_id",	"pole2_id",	"pole3_id",	"pole4_id",	"pole5_id",
            "ant1",	"ant2",	"ant3",	"ant4",	"ant5",	"ant1_donor",	"ant2_donor",	"ant3_donor",	"ant4_donor",	"ant5_donor",
            "br1",	"br2",	"br3",	"br4",	"br5",	"Ant1_Port",	"Ant2_Port",	"Ant3_Port",	"Ant4_Port",	"Ant5_Port",
            "Ant1_Height",	"Ant2_Height",	"Ant3_Height",	"Ant4_Height",	"Ant5_Height",
            "mdt1",	"mdt2",	"mdt3",	"mdt4",	"mdt5",	"edt1",	"edt2",	"edt3",	"edt4",	"edt5",
            "sp_loss1",	"sp_loss2",	"sp_loss3",	"sp_loss4",	"sp_loss5",
            "fne_id",	"tx_rx_pa0",	"tx_rx_pa1",	"tx_rx_pa2",	"tx_rx_pa3",	"tx_rx_pa4",	"tx_rx_pa5",
            "cdma_cell",	"g_cell_pa0",	"g_cell_pa1",	"g_trx_pa0",	"g_trx_pa1",	"u_cell_pa0",	"u_cell_pa1",	"u_cell_pa2",	"u_cell_pa3",
            "l_cell_pa0",	"l_cell_pa1",	"l_cell_pa2",	"l_cell_pa3",	"l_cell_pa4",	"l_cell_pa5",
            "WiFi_Cell_code",	"otsr",	"bbu_name",	"vam",	"remark",
            "n_cell_pa0",	"n_cell_pa1",	"n_cell_pa2",	"n_cell_pa3",	"n_cell_pa4",	"n_cell_pa5",	"n_cell_pa6",	"n_cell_pa7",
            "tx_rx_pa6",	"tx_rx_pa7"};


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE,
                Manifest.permission.READ_EXTERNAL_STORAGE}, PackageManager.PERMISSION_GRANTED);
        if (Build.VERSION.SDK_INT >= 30){
            if (!Environment.isExternalStorageManager()){
                Intent getpermission = new Intent();
                getpermission.setAction( Settings.ACTION_MANAGE_ALL_FILES_ACCESS_PERMISSION);
                startActivity(getpermission);
            }
        }
        edt_input_site_code = findViewById(R.id.edt_Site_code);
    }

    public void buttonCreateExcel(View view){
        int size = strings.length;
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        HSSFSheet hssfSheet = hssfWorkbook.createSheet();

        HSSFRow hssfRow = hssfSheet.createRow(1);
        HSSFCell hssfCell = hssfRow.createCell(0);

        hssfCell.setCellValue(edt_input_site_code.getText().toString());

        for(int rownum = 0; rownum <1; rownum++){
            Row row = hssfSheet.createRow(rownum);
            for(int cellnum = 0; cellnum<size;cellnum++){
                Cell cell = row.createCell(cellnum);
                cell.setCellValue(strings[cellnum]);
            }
        }

            try {
                if (!filePath.exists()) {
                    filePath.createNewFile();
                } else {
                }
                FileOutputStream fileOutputStream = new FileOutputStream(filePath);
                hssfWorkbook.write(fileOutputStream);
                Toast.makeText(this, "table one has created!!", Toast.LENGTH_LONG).show();

                if (fileOutputStream != null) {
                    fileOutputStream.flush();
                    fileOutputStream.close();
                } else {
                }
            } catch (Exception e) {
            }
        }
}