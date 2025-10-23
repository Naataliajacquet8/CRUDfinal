using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace CRUDfinal
{
    public partial class frmUsuarios : Form
    {
        
        private string conexionAccess = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source= C:\Users\Admin\Documents\BDfinal1.accdb";
        private OleDbConnection cn;
        List<Usuarios> usuario = new List<Usuarios>();
        public frmUsuarios()
        {
            InitializeComponent();
            this.CenterToScreen();
            
        }

        private void frmUsuarios_Load(object sender, EventArgs e)
        {
            actualizarGrilla();
            
        }

        void conectar()
        {
            try
            {
                cn = new OleDbConnection(conexionAccess);
                cn.Open();
            }
            catch (Exception ex)
            {
                throw new Exception("Error al conectar la base de datos: " + ex.Message);
            }

        }

        void desconectar()
        {

            try
            {
                if (cn != null && cn.State == ConnectionState.Open)
                {
                    cn.Close();
                    cn.Dispose();
                }


            }
            catch (Exception ex)
            {
                throw new Exception("Error al desconectar la base de datos: " + ex.Message);
            }
        }

        List<Usuarios> buscarPersonas()
        {
            usuario.Clear();
            try
            {
                conectar();
                string sql = "Select * From Usuarios ORDER BY Id_usuario ASC";
                
                using (OleDbCommand cmd = new OleDbCommand(sql, cn))
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        usuario.Add(new Usuarios
                        {
                            Id_usuario= Convert.ToInt32(reader["Id_usuario"]),
                            Apellido = reader["Apellido"].ToString() ?? string.Empty,
                            Nombre = reader["Nombre"].ToString() ?? string.Empty,
                            DNI = Convert.ToInt32(reader["DNI"]),
                            Telefono= Convert.ToInt32(reader["Telefono"]),
                            Direccion = reader["Direccion"].ToString() ?? string.Empty
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error al buscar personas: " + ex.Message);
            }
            finally
            {
                desconectar();
            }
            return usuario;
        }

        void actualizarGrilla()
        {

            bloquearControles(gbDatosPersonales);
            dgvUsuarios.DataSource = null;
            buscarPersonas();
            dgvUsuarios.DataSource = usuario;
            dgvUsuarios.Columns["Id_usuario"].Visible = false;

            cmbBuscar.Items.Clear();
            foreach (var u in usuario)
            {
                cmbBuscar.Items.Add(u.Nombre);
            }

        }

        void limpiarGrilla(Control control)
        {
            foreach (Control item in control.Controls)
            {
                if (item is TextBox txt) txt.Text = null;
                //if (item is CheckBox chk) chk.Checked = false;
                //if (item is GroupBox || item is Panel) limpiarGrilla((Control)item);
            }
        }

        private void dgvUsuarios_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvUsuarios.SelectedRows.Count > 0)
            {
                Usuarios seleccionado = (Usuarios)dgvUsuarios.SelectedRows[0].DataBoundItem;
                txtNombre.Text = seleccionado.Nombre;
                txtApellido.Text = seleccionado.Apellido;
                txtDNI.Text = seleccionado.DNI.ToString();
                txtTelefono.Text = seleccionado.Telefono.ToString();
                txtDireccion.Text = seleccionado.Direccion;
            }
        }

        int indiceSeleccionado = -1; // variable de clase 
        private void btnModificar_Click(object sender, EventArgs e)
        {
            if (btnModificar.Text == "Modificar")
            {
                if (dgvUsuarios.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Seleccione un usuario para modificar.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                indiceSeleccionado = dgvUsuarios.SelectedRows[0].Index;

                desbloquearControles(gbDatosPersonales);
                txtDNI.Enabled = false;

                btnModificar.Text = "Guardar";
                btnAgregar.Visible = false;
                btnEliminar.Text = "Cancelar";

                txtNombre.Focus();
            }
            else
            {
                if (indiceSeleccionado < 0 || indiceSeleccionado >= usuario.Count)
                {
                    MessageBox.Show("No se encontró el usuario a modificar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Actualizar el usuario en la lista
                
                usuario[indiceSeleccionado].Nombre = txtNombre.Text;
                usuario[indiceSeleccionado].Apellido = txtApellido.Text;
                usuario[indiceSeleccionado].Telefono = Convert.ToInt32(txtTelefono.Text);
                usuario[indiceSeleccionado].DNI = Convert.ToInt32(txtDNI.Text);
                usuario[indiceSeleccionado].Direccion = txtDireccion.Text;

                // Guardar cambios en la base de datos
                modificarUsuario(
                    usuario[indiceSeleccionado].Nombre = txtNombre.Text,
                    usuario[indiceSeleccionado].Apellido = txtApellido.Text,
                    usuario[indiceSeleccionado].Telefono = Convert.ToInt32(txtTelefono.Text),
                    usuario[indiceSeleccionado].DNI = Convert.ToInt32(txtDNI.Text),
                    usuario[indiceSeleccionado].Direccion = txtDireccion.Text
                );

                MessageBox.Show("Se modificó correctamente el usuario", "Modificar usuario", MessageBoxButtons.OK, MessageBoxIcon.Information);

                btnModificar.Text = "Modificar";
                btnAgregar.Visible = true;
                btnEliminar.Text = "Eliminar";
                actualizarGrilla();
                limpiarGrilla(gbDatosPersonales);
                dgvUsuarios.Focus();
            }

        
        }

        void modificarUsuario(string nombre, string apellido, int telefono, int dni, string direccion)
        {
            try
            {
                conectar();

                string sql = "UPDATE Usuarios SET Nombre = ?, Apellido = ?, Telefono = ?, Direccion = ? WHERE DNI=?";

                using (OleDbCommand cmd = new OleDbCommand(sql, cn))
                {
                    cmd.Parameters.AddWithValue("?", nombre);
                    cmd.Parameters.AddWithValue("?", apellido);
                    cmd.Parameters.AddWithValue("?", telefono);
                    cmd.Parameters.AddWithValue("?", direccion);
                    cmd.Parameters.AddWithValue("?", dni);

                    int filasAfectadas = cmd.ExecuteNonQuery();

                    if (filasAfectadas == 0)
                    {
                        MessageBox.Show("No se encontró ningún usuario con ese DNI para modificar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al modificar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                desconectar();
            }
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (btnAgregar.Text == "Agregar")
            {
                btnAgregar.Text = "Guardar";
                btnModificar.Visible = false;
                btnEliminar.Text = "Cancelar";


                desbloquearControles(gbDatosPersonales);
                
                limpiarGrilla(gbDatosPersonales);
                txtNombre.Focus();
            }
            else
            {
                // Validar campos vacíos
                if (string.IsNullOrWhiteSpace(txtNombre.Text) || string.IsNullOrWhiteSpace(txtApellido.Text) ||
                    string.IsNullOrWhiteSpace(txtTelefono.Text) || string.IsNullOrWhiteSpace(txtDNI.Text))
                {
                    MessageBox.Show("Complete todos los campos antes de continuar.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Validar formato de teléfono
                if (!Regex.IsMatch(txtTelefono.Text, @"^11\d{8}$"))
                {
                    MessageBox.Show("El teléfono debe comenzar con 11 y tener 10 dígitos en total.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Validar formato de DNI
                if (!Regex.IsMatch(txtDNI.Text, @"^\d{8}$"))
                {
                    MessageBox.Show("El DNI debe tener exactamente 8 dígitos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                btnAgregar.Text = "Agregar";
                btnEliminar.Text = "Eliminar";
                btnModificar.Visible = true;

                // Verifica que el DNI no esta en uso
                if (ExisteDNI(txtDNI.Text.Trim()))
                {
                    MessageBox.Show("Ya existe un usuario con ese DNI.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                bloquearControles(this);

                // Crear nuevo usuario
                Usuarios nuevoUsuario = new Usuarios()
                {
                    Id_usuario = (usuario.Any()) ? usuario.Max(x => x.Id_usuario) + 1 : 1,
                    Nombre = txtNombre.Text,
                    Apellido = txtApellido.Text,
                    DNI = Convert.ToInt32(txtDNI.Text),
                    Telefono = Convert.ToInt32(txtTelefono.Text),
                    Direccion = txtDireccion.Text
                };

                usuario.Add(nuevoUsuario);
                insertarUsuario(txtNombre.Text, txtApellido.Text,Convert.ToInt32(txtTelefono.Text), Convert.ToInt32(txtDNI.Text),txtDireccion.Text);
                MessageBox.Show("Se inserto correctamente el nuevo usuario", "Insertar usuario", MessageBoxButtons.OK, MessageBoxIcon.Information);
                actualizarGrilla(); // Esto actualiza DataGridView y ComboBox
                limpiarGrilla(gbDatosPersonales);
                dgvUsuarios.Focus();
            }
        }

        private bool ExisteDNI(string dni)
        {
            try
            {
                conectar();
                string sql = "SELECT COUNT(*) FROM Usuarios WHERE DNI = ?";
                using (OleDbCommand cmd = new OleDbCommand(sql, cn))
                {
                    cmd.Parameters.AddWithValue("?", dni);
                    int count = (int)cmd.ExecuteScalar();
                    return count > 0; // true si ya existe
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al verificar el DNI: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true; // Devuelvo true para evitar que continúe si algo sale mal
            }
            finally
            {
                desconectar();
            }
        }
        private void insertarUsuario(string nombre, string apellido, int telefono, int dni, string direccion)
        {
            
            try
            {
                conectar();
                string sql = "INSERT INTO Usuarios (Nombre, Apellido, Telefono, DNI, Direccion) VALUES (?,?,?,?,?)";
                
                using (OleDbCommand cmd =new OleDbCommand (sql, cn))
                {
                    
                    cmd.Parameters.AddWithValue("?", nombre);
                    cmd.Parameters.AddWithValue("?", apellido);
                    cmd.Parameters.AddWithValue("?", telefono);
                    cmd.Parameters.AddWithValue("?",dni);
                    cmd.Parameters.AddWithValue("?", direccion);

                    cmd.ExecuteNonQuery();
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Error al insertar" + ex.Message);
            }
            finally
            {
                desconectar();
            }
        }

        private void bloquearControles(Control control)
        {
            foreach (Control item in control.Controls)
            {
                if (item is TextBox txt) txt.Enabled = false;
                
            }
        }

        private void desbloquearControles(Control contrl)
        {
            foreach(Control item in contrl.Controls)
            {
                if (item is TextBox txt) txt.Enabled = true;
            }
        }

        private void cmbBuscar_SelectedIndexChanged(object sender, EventArgs e)
        {
            string nombreSeleccionado = cmbBuscar.SelectedItem.ToString();

            Usuarios seleccinado = usuario.FirstOrDefault(u => u.Nombre == nombreSeleccionado);
            if (seleccinado != null)
            {
                
                txtNombre.Text = seleccinado.Nombre;
                txtApellido.Text = seleccinado.Apellido;
                txtDNI.Text = seleccinado.DNI.ToString();
                txtDireccion.Text = seleccinado.Direccion;
                txtTelefono.Text = seleccinado.Telefono.ToString();


            }
        }

         private void btnEliminar_Click_1(object sender, EventArgs e)
 {
     if (btnEliminar.Text == "Eliminar")
     {
         // Verificar selección
         if (dgvUsuarios.CurrentRow == null)
         {
             MessageBox.Show("Seleccioná un usuario para eliminar.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
             return;
         }

         // Confirmar
         DialogResult respuesta = MessageBox.Show(
             "¿Estás segura de que querés eliminar este usuario?",
             "Confirmar eliminación",
             MessageBoxButtons.YesNo,
             MessageBoxIcon.Question
         );

         if (respuesta == DialogResult.No)
             return;

         try
         {
             // Obtener el ID del usuario seleccionado
             int idUsuario = Convert.ToInt32(dgvUsuarios.CurrentRow.Cells["Id_usuario"].Value);

             // ELIMINAR DE LA BASE DE DATOS
             conectar();
             string sql = "DELETE FROM Usuarios WHERE Id_usuario = ?";
             using (OleDbCommand cmd = new OleDbCommand(sql, cn))
             {
                 cmd.Parameters.AddWithValue("?", idUsuario);
                 int filas = cmd.ExecuteNonQuery();

                 if (filas > 0)
                 {
                     MessageBox.Show("Usuario eliminado correctamente de la base de datos.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 }
                 else
                 {
                     MessageBox.Show("No se encontró el usuario en la base de datos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 }
             }

             //ELIMINAR TAMBIÉN DE LA LISTA
             Usuarios eliminado = usuario.FirstOrDefault(u => u.Id_usuario == idUsuario);
             if (eliminado != null)
             {
                 usuario.Remove(eliminado);
                 actualizarGrilla();
             }
         }
         catch (Exception ex)
         {
             MessageBox.Show("Error al eliminar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
         }
         finally
         {
             desconectar();
         }
     }
     else // modo "Cancelar"
     {
         btnAgregar.Text = "Agregar";
         btnEliminar.Text = "Eliminar";
         btnModificar.Text = "Modificar";
         btnModificar.Visible = true;
         btnAgregar.Visible = true;
         limpiarGrilla(gbDatosPersonales);
     }
 }
    

    }
}
