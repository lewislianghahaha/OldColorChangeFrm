using System;
using System.Data;
using System.Threading;
using System.Windows.Forms;
using Mergedt;
using OldColorChangeFrm.Logic;

namespace OldColorChangeFrm
{
    public partial class Main : Form
    {
        Load load=new Load();
        TaskLogic task=new TaskLogic();

        public Main()
        {
            InitializeComponent();
            OnRegisterEvents();
            OnShow();
        }

        private void OnRegisterEvents()
        {
            tmclose.Click += Tmclose_Click;
            btngen.Click += Btngen_Click;
            comlist.SelectedIndexChanged += Comlist_SelectedIndexChanged;
        }

        private void OnShow()
        {
            //初始化下拉列表
            OnShowStatusList();
        }

        private void Comlist_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //获取下拉列表所选值
                var dvordertylelist = (DataRowView)comlist.Items[comlist.SelectedIndex];
                var typeId = Convert.ToInt32(dvordertylelist["Id"]);
                GlobalClasscs.ChooseType.ChooseTypeId = typeId;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btngen_Click(object sender, EventArgs e)
        {
            try
            {
                task.TaskId = 0;

                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();

                if(!task.ResultMark) throw new Exception("运算不成功,请联系管理员");
                else
                {
                    var clickMessage = "运算成功,是否进行导出";
                    if (MessageBox.Show(clickMessage, "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        Exportdt();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 关闭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Tmclose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        ///子线程使用(重:用于监视功能调用情况,当完成时进行关闭LoadForm)
        /// </summary>
        private void Start()
        {
            task.StartTask();

            //当完成后将Form2子窗体关闭
            this.Invoke((ThreadStart)(() => {
                load.Close();
            }));
        }

        /// <summary>
        /// 导出
        /// </summary>
        void Exportdt()
        {

            try
            {
                var saveFileDialog = new SaveFileDialog { Filter = "Xlsx文件|*.xlsx" };
                if (saveFileDialog.ShowDialog() != DialogResult.OK) return;
                var fileAdd = saveFileDialog.FileName;

                task.TaskId = 1;
                task.FileAddress = fileAdd;

                //使用子线程工作(作用:通过调用子线程进行控制Load窗体的关闭情况)
                new Thread(Start).Start();
                load.StartPosition = FormStartPosition.CenterScreen;
                load.ShowDialog();

                if (!task.ResultMark) throw new Exception("导出异常");
                else
                {
                    MessageBox.Show($"导出成功!可从EXCEL中查阅导出效果", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 初始化单据状态下拉列表
        /// </summary>
        private void OnShowStatusList()
        {
            var dt = new DataTable();

            //创建表头
            for (var i = 0; i < 2; i++)
            {
                var dc = new DataColumn();
                switch (i)
                {
                    case 0:
                        dc.ColumnName = "Id";
                        break;
                    case 1:
                        dc.ColumnName = "Name";
                        break;
                }
                dt.Columns.Add(dc);
            }

            //创建行内容
            for (var j = 0; j < 2; j++)
            {
                var dr = dt.NewRow();

                switch (j)
                {
                    case 0:
                        dr[0] = "0";
                        dr[1] = "以横向方式导出";
                        break;
                    case 1:
                        dr[0] = "1";
                        dr[1] = "以竖向方式导出";
                        break;
                }
                dt.Rows.Add(dr);
            }

            comlist.DataSource = dt;
            comlist.DisplayMember = "Name"; //设置显示值
            comlist.ValueMember = "Id";    //设置默认值内码
        }

    }
}
