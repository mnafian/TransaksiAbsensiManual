Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types

'Form2 untuk crystal report cetak data berdasarkan ID
'Form3 untuk crystal report cetak semua data

Public Class Form1
    Dim datable As New DataTable
    Dim id_tran As String

    'Load Id karyawan untuk combobox
    Sub LoadIdKaryawan()
        cmd = New OracleCommand("select ID_TRAN from T_ABSEN_MANUAL order by ID_TRAN", con)
        dr = cmd.ExecuteReader
        While dr.Read
            ComboBox1.Items.Add(dr("ID_TRAN"))
        End While
    End Sub

    'Load data absensi manual ke datagridview 
    Sub LoadDataKaryawan()
        id_tran = ComboBox1.SelectedItem.ToString
        da = New OracleDataAdapter(" select b.npk,c.nama,b.kerja_1 as jam_masuk, b.kerja_2 as jam_pulang, b.lembur_1 as lembur_mulai," &
                                   " b.lembur_2 as lembur_pulang, b.ijin_1 as ijin_keluar, b.ijin_2 as ijin_kembali, d.n_jabatan as jabatan," &
                                   " e.nama_departemen as departemen, a.ket as keterangan, b.terlambat, b.pul_awal as pulang_awal ," &
                                   " b.jum_jamker as jam_kerja from t_absen_manual a" &
                                   " join t_absen_manual_detil b on a.id_tran=b.id_tran" &
                                   " join karyawan c on b.npk=c.npk join jab d on c.jab=d.id_jabatan" &
                                   " join dept e on c.dept=e.id_departemen where a.id_tran='" + id_tran + "' order by a.id_tran", con)
        ds = New DataSet
        da.Fill(ds, "t_absen")
        DataGridView1.DataSource = ds.Tables("t_absen")
        DataGridView1.ReadOnly = True
    End Sub

    'Cetak Data berdasarkan Id
    Sub CetakDataAbsen()
        id_tran = ComboBox1.SelectedItem.ToString
        da = New OracleDataAdapter("select a.id_tran,a.status,to_char(a.dt_tran, 'fmdd MON yyyy')as tanggal,b.npk,c.nama,b.kerja_1 as jam_masuk, b.kerja_2 as jam_pulang, b.lembur_1 as lembur_mulai, b.lembur_2 as lembur_pulang, b.ijin_1 as ijin_keluar, b.ijin_2 as ijin_kembali, d.n_jabatan as jabatan, e.nama_departemen as departemen, a.ket as keterangan, b.terlambat, b.pul_awal as pulang_awal ,b.jum_jamker as jam_kerja,f.nama as pembuat, j.n_jabatan as jab_buat, g.nama as pemeriksa, k.n_jabatan as jab_periksa,h.nama as menyetujui, l.n_jabatan as jab_setuju,i.nama as mengetahui, m.n_jabatan as jab_mengetahui from t_absen_manual a left join t_absen_manual_detil b on a.id_tran=b.id_tran left join karyawan c on b.npk=c.npk left join jab d on c.jab=d.id_jabatan left join dept e on c.dept=e.id_departemen left join karyawan f on f.npk=a.lev_1 left join karyawan g on g.npk=a.lev_2 left join karyawan h on h.npk=a.lev_3 left join karyawan i on i.npk=a.lev_4 left join jab j on j.id_jabatan=f.jab left join jab k on k.id_jabatan=g.jab left join jab l on l.id_jabatan=h.jab left join jab m on m.id_jabatan=i.jab where a.id_tran='" + id_tran + "' order by a.id_tran", con)
        ds = New DataSet
        da.Fill(datable)
        da.Fill(ds, "t_absen")

        If datable.Rows.Count = 0 Then
            Label4.Text = "-"
            Label6.Text = "-"
            Label8.Text = "-"
            Label12.Text = "-"
            Label13.Text = "-"
            Label14.Text = "-"
            Label15.Text = "-"
        Else
            Label4.Text = datable.Rows(0)("status").ToString()
            Label6.Text = datable.Rows(0)("tanggal").ToString()
            Label8.Text = datable.Rows(0)("keterangan").ToString()
            Label12.Text = datable.Rows(0)("pembuat").ToString()
            Label13.Text = datable.Rows(0)("pemeriksa").ToString()
            Label14.Text = datable.Rows(0)("menyetujui").ToString()
            Label15.Text = datable.Rows(0)("mengetahui").ToString()
        End If
        datable.Clear()
    End Sub

    'Cetak semua laporan transaksi manual
    Sub CetakDataAbsenAll()
        id_tran = ComboBox1.SelectedItem.ToString
        da = New OracleDataAdapter("select a.id_tran,a.status,to_char(a.dt_tran, 'fmdd MON yyyy')as tanggal,b.npk,c.nama,b.kerja_1 as jam_masuk, b.kerja_2 as jam_pulang, b.lembur_1 as lembur_mulai, b.lembur_2 as lembur_pulang, b.ijin_1 as ijin_keluar, b.ijin_2 as ijin_kembali, d.n_jabatan as jabatan, e.nama_departemen as departemen, a.ket as keterangan, b.terlambat, b.pul_awal as pulang_awal ,b.jum_jamker as jam_kerja,f.nama as pembuat, j.n_jabatan as jab_buat, g.nama as pemeriksa, k.n_jabatan as jab_periksa,h.nama as menyetujui, l.n_jabatan as jab_setuju,i.nama as mengetahui, m.n_jabatan as jab_mengetahui from t_absen_manual a left join t_absen_manual_detil b on a.id_tran=b.id_tran left join karyawan c on b.npk=c.npk left join jab d on c.jab=d.id_jabatan left join dept e on c.dept=e.id_departemen left join karyawan f on f.npk=a.lev_1 left join karyawan g on g.npk=a.lev_2 left join karyawan h on h.npk=a.lev_3 left join karyawan i on i.npk=a.lev_4 left join jab j on j.id_jabatan=f.jab left join jab k on k.id_jabatan=g.jab left join jab l on l.id_jabatan=h.jab left join jab m on m.id_jabatan=i.jab", con)
        ds = New DataSet
        da.Fill(ds, "t_absen")
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call LoadIdKaryawan()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call koneksi()
        Call LoadDataKaryawan()
        Call CetakDataAbsen()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call koneksi()
        id_tran = ComboBox1.SelectedItem.ToString
        If id_tran = "All" Then
            Call CetakDataAbsenAll()
            Form3.Show()
        Else
            Form2.Show()
        End If
    End Sub
End Class
