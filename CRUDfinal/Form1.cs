using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
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
                string sql = "Select * From Usuarios";
                
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
                            Direccion = reader["Dirección"].ToString() ?? string.Empty
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
            // foreach (Control item in control.Controls)
            // {
             //  if (item is TextBox txt) txt.Text = string.Empty;
            //   if (item is ComboBox cmb) cmb.SelectedIndex = -1;
            // }
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

        private void btnModificar_Click(object sender, EventArgs e)
        {
            if (btnModificar.Text == "Modificar")
            {
                btnModificar.Text = "Guardar Cambios";
                btnAgregar.Visible = false;
                btnEliminar.Text = "Cancelar";

                
                txtNombre.Focus();
            }
            else
            {
                btnModificar.Text = "Modificar";
                btnAgregar.Visible = true;
                btnEliminar.Text = "Eliminar";
                

                
                int indice = dgvUsuarios.SelectedRows[0].Index;

                // Actualizar usuario existente
                usuario[indice].Nombre = txtNombre.Text;
                usuario[indice].Apellido = txtApellido.Text;
                usuario[indice].DNI = Convert.ToInt32(txtDNI.Text);
                usuario[indice].Telefono = Convert.ToInt32(txtTelefono.Text);
                usuario[indice].Direccion = txtDireccion.Text;

                actualizarGrilla();
                limpiarGrilla(gbDatosPersonales);
                dgvUsuarios.Focus();
            }
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            if (btnAgregar.Text == "Agregar")
            {
                btnAgregar.Text = "Guardar";
                btnModificar.Visible = false;
                btnEliminar.Text = "Cancelar";

                
                limpiarGrilla(gbDatosPersonales);
                txtNombre.Focus();
            }
            else
            {
                btnAgregar.Text = "Agregar";
                btnEliminar.Text = "Eliminar";
                btnModificar.Visible = true;

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

        private void insertarUsuario(string nombre, string apellido, int telefono, int dni, string direccion)
        {
            try
            {
                conectar();
                string sql = "INSERT INTO Usuarios (Nombre, Apellido, Telefono, DNI, Dirección) VALUES (?,?,?,?,?)";
                
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

        private void btnEliminar_Click(object sender, EventArgs e)
        {
           
            if (btnEliminar.Text == "Eliminar")
            {
                // Verificar que haya algo seleccionado
                if (dgvUsuarios.CurrentRow == null)
                {
                    MessageBox.Show("Seleccioná un usuario para eliminar.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
        
                // Confirmar antes de eliminar
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
                    // Obtener el Id del usuario seleccionado
                    int idUsuario = Convert.ToInt32(dgvUsuarios.CurrentRow.Cells["Id_Usuario"].Value);
        
                    // Buscar el usuario en la lista
                    Usuarios eliminado = usuario.FirstOrDefault(u => u.Id_Usuario == idUsuario);
        
                    if (eliminado != null)
                    {
                        usuario.Remove(eliminado); // eliminar de la lista
                        actualizarGrilla();        // refrescar grilla
                        MessageBox.Show("Usuario eliminado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("No se encontró el usuario seleccionado en la lista.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al eliminar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else // si está en modo Cancelar
            {
                btnAgregar.Text = "Agregar";
                btnEliminar.Text = "Eliminar";
                btnModificar.Text = "Modificar";
                btnModificar.Visible = true;
                btnAgregar.Visible = true;
        
                limpiarGrilla(this); // limpia todos los TextBox
            }
               
        }


    }
}
