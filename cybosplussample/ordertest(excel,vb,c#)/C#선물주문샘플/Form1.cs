using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace WindowsApplication1
{
	/// <summary>
	/// Form1에 대한 요약 설명입니다.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		/// <summary>
		/// 필수 디자이너 변수입니다.
		/// </summary>
		private System.ComponentModel.Container components = null;
        
		public Form1()
		{
			//
			// Windows Form 디자이너 지원에 필요합니다.
			//
			InitializeComponent();

			//
			// TODO: InitializeComponent를 호출한 다음 생성자 코드를 추가합니다.
			//
		}

		/// <summary>
		/// 사용 중인 모든 리소스를 정리합니다.
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

		#region Windows Form 디자이너에서 생성한 코드
		/// <summary>
		/// 디자이너 지원에 필요한 메서드입니다.
		/// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
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
            this.button1.Text = "선물주문(CpTd6831)";
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
		/// 해당 응용 프로그램의 주 진입점입니다.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
            //주문 초기화
            int nRet;
            do
            {
                nRet = m_TdUtil.TradeInit(0);
            } while (nRet != 0);
            
            //주문이벤트등록
            m_CpFConclusion.Received += new DSCBO1Lib._IDibEvents_ReceivedEventHandler(CpFConclusion_OnReceived);
            m_CpFConclusion.Subscribe();

            //계좌번호 얻기                        
            Console.WriteLine(m_TdUtil.AccountNumber.ToString());
            m_arAccount = (Array)m_TdUtil.AccountNumber;
            Console.WriteLine(m_arAccount.GetValue(0));                       

            //FutureCurr Test
            m_FutureCurr.SetInputValue(0, m_FCode.GetData(0, 0)); //최근월물 코드
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
            //최근월물의 현재가격으로 선물주문 요청
            m_FutureMst.SetInputValue(0, m_FCode.GetData(0, 0)); // 종목코드
            m_FutureMst.BlockRequest();

            //주문 설정
            m_6831.SetInputValue(0, "1"); // 선물/옵션 구분 ("1":선물, "2":옵션, "3":개별주식옵션...)            
            m_6831.SetInputValue(1, m_arAccount.GetValue(0)); // 계좌번호
            m_6831.SetInputValue(2, m_FCode.GetData(0,0)); // 종목코드
            m_6831.SetInputValue(3, 1); // 주문수량
            m_6831.SetInputValue(4, m_FutureMst.GetHeaderValue (71)); //71은 현재가
            m_6831.SetInputValue(5, "2");
            m_6831.SetInputValue(6, "1");
            m_6831.SetInputValue(7, "0");
            int nRet = m_6831.BlockRequest();
            Console.WriteLine(m_6831.GetDibMsg1()); 
        } 
	}
}
