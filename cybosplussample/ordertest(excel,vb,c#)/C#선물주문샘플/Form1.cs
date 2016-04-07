using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace WindowsApplication1
{
	/// <summary>
	/// Form1�� ���� ��� �����Դϴ�.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		/// <summary>
		/// �ʼ� �����̳� �����Դϴ�.
		/// </summary>
		private System.ComponentModel.Container components = null;
        
		public Form1()
		{
			//
			// Windows Form �����̳� ������ �ʿ��մϴ�.
			//
			InitializeComponent();

			//
			// TODO: InitializeComponent�� ȣ���� ���� ������ �ڵ带 �߰��մϴ�.
			//
		}

		/// <summary>
		/// ��� ���� ��� ���ҽ��� �����մϴ�.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form �����̳ʿ��� ������ �ڵ�
		/// <summary>
		/// �����̳� ������ �ʿ��� �޼����Դϴ�.
		/// �� �޼����� ������ �ڵ� ������� �������� ���ʽÿ�.
		/// </summary>
		private void InitializeComponent()
		{
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(77, 109);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(133, 83);
            this.button1.TabIndex = 0;
            this.button1.Text = "�����ֹ�(CpTd6831)";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
            this.ClientSize = new System.Drawing.Size(292, 266);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// �ش� ���� ���α׷��� �� �������Դϴ�.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
            //�ֹ� �ʱ�ȭ
            int nRet;
            do
            {
                nRet = m_TdUtil.TradeInit(0);
            } while (nRet != 0);
            
            //�ֹ��̺�Ʈ���
            m_CpFConclusion.Received += new DSCBO1Lib._IDibEvents_ReceivedEventHandler(CpFConclusion_OnReceived);
            m_CpFConclusion.Subscribe();

            //���¹�ȣ ���                        
            Console.WriteLine(m_TdUtil.AccountNumber.ToString());
            m_arAccount = (Array)m_TdUtil.AccountNumber;
            Console.WriteLine(m_arAccount.GetValue(0));                       

            //FutureCurr Test
            m_FutureCurr.SetInputValue(0, m_FCode.GetData(0, 0)); //�ֱٿ��� �ڵ�
            m_FutureCurr.Received += new DSCBO1Lib._IDibEvents_ReceivedEventHandler(m_FutureCurr_OnReceived);
            m_FutureCurr.Subscribe();            
              
		}

        void m_FutureCurr_OnReceived()
        {
            Console.WriteLine("m_FutureCurr_OnReceived");     
        }

        void CpFConclusion_OnReceived()
        {
            Console.WriteLine("CpFConclusion_OnReceived");     
        }
        
        //*********************************************************
        DSCBO1Lib.CpFConclusionClass m_CpFConclusion = new DSCBO1Lib.CpFConclusionClass();
        DSCBO1Lib.FutureCurrClass m_FutureCurr = new DSCBO1Lib.FutureCurrClass();
        DSCBO1Lib.FutureMst m_FutureMst = new DSCBO1Lib.FutureMst();

        CPTRADELib.CpTdUtilClass m_TdUtil = new CPTRADELib.CpTdUtilClass();
        CPTRADELib.CpTd6831Class m_6831 = new CPTRADELib.CpTd6831Class();
        CPUTILLib.CpFutureCode m_FCode = new CPUTILLib.CpFutureCode();       
        Array m_arAccount;
        //*********************************************************

        private Button button1;
        private void button1_Click(object sender, EventArgs e)
        {
            //�ֱٿ����� ���簡������ �����ֹ� ��û
            m_FutureMst.SetInputValue(0, m_FCode.GetData(0, 0)); // �����ڵ�
            m_FutureMst.BlockRequest();

            //�ֹ� ����
            m_6831.SetInputValue(0, "1"); // ����/�ɼ� ���� ("1":����, "2":�ɼ�, "3":�����ֽĿɼ�...)            
            m_6831.SetInputValue(1, m_arAccount.GetValue(0)); // ���¹�ȣ
            m_6831.SetInputValue(2, m_FCode.GetData(0,0)); // �����ڵ�
            m_6831.SetInputValue(3, 1); // �ֹ�����
            m_6831.SetInputValue(4, m_FutureMst.GetHeaderValue (71)); //71�� ���簡
            m_6831.SetInputValue(5, "2");
            m_6831.SetInputValue(6, "1");
            m_6831.SetInputValue(7, "0");
            int nRet = m_6831.BlockRequest();
            Console.WriteLine(m_6831.GetDibMsg1()); 
        } 
	}
}
