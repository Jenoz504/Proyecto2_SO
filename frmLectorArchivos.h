#pragma once

namespace Proyecto2_SO {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::IO;

	/// <summary>
	/// Resumen de frmLectorArchivos
	/// </summary>
	public ref class frmLectorArchivos : public System::Windows::Forms::Form
	{
	public:
		frmLectorArchivos(void)
		{
			InitializeComponent();
			//
			//TODO: agregar código de constructor aquí
			//
		}

	protected:
		/// <summary>
		/// Limpiar los recursos que se estén usando.
		/// </summary>
		~frmLectorArchivos()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^ btnExplorador;
	protected:
	private: System::Windows::Forms::Label^ lblDireccion;

	private:
		/// <summary>
		/// Variable del diseñador necesaria.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido de este método con el editor de código.
		/// </summary>
		void InitializeComponent(void)
		{
			this->btnExplorador = (gcnew System::Windows::Forms::Button());
			this->lblDireccion = (gcnew System::Windows::Forms::Label());
			this->SuspendLayout();
			// 
			// btnExplorador
			// 
			this->btnExplorador->Location = System::Drawing::Point(288, 59);
			this->btnExplorador->Name = L"btnExplorador";
			this->btnExplorador->Size = System::Drawing::Size(99, 45);
			this->btnExplorador->TabIndex = 0;
			this->btnExplorador->Text = L"Abrir Plantilla de excel";
			this->btnExplorador->UseVisualStyleBackColor = true;
			this->btnExplorador->Click += gcnew System::EventHandler(this, &frmLectorArchivos::btnExplorador_Click);
			// 
			// lblDireccion
			// 
			this->lblDireccion->AutoSize = true;
			this->lblDireccion->Location = System::Drawing::Point(12, 144);
			this->lblDireccion->Name = L"lblDireccion";
			this->lblDireccion->Size = System::Drawing::Size(10, 13);
			this->lblDireccion->TabIndex = 1;
			this->lblDireccion->Text = L"-";
			// 
			// frmLectorArchivos
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(601, 327);
			this->Controls->Add(this->lblDireccion);
			this->Controls->Add(this->btnExplorador);
			this->Name = L"frmLectorArchivos";
			this->Text = L"frmLectorArchivos";
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void btnExplorador_Click(System::Object^ sender, System::EventArgs^ e) {
		String^ LastFileDirectory;
		String^ Filename;

		//If last directory is not valid then default to My Documents (if you don't include this the catch below won't occur for null strings so the start directory is undefined)
		if (!Directory::Exists(LastFileDirectory))
			LastFileDirectory = Environment::GetFolderPath(Environment::SpecialFolder::MyDocuments);

		//----- FILE OPEN DIALOG BOX -----
		OpenFileDialog^ SelectFileDialog = gcnew OpenFileDialog();

		SelectFileDialog->Filter = "All Files (*.*)|*.*";		//"txt files (*.txt)|*.txt|All files (*.*)|*.*"
		SelectFileDialog->FilterIndex = 1;				//(First entry is 1, not 0)
		try
		{
			SelectFileDialog->InitialDirectory = LastFileDirectory;
		}
		catch (Exception^)
		{
			SelectFileDialog->InitialDirectory = Environment::GetFolderPath(Environment::SpecialFolder::MyDocuments);
		}
		//Display dialog box
		if (SelectFileDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK)
		{
			//----- OPEN THE FILE -----
			Filename = SelectFileDialog->FileName;
			lblDireccion->Text = Filename;
			//STORE LAST USED DIRECTORY
			if (Filename->LastIndexOf("\\") >= 0)
				LastFileDirectory = Filename->Substring(0, (Filename->LastIndexOf("\\") + 1));
		}
	}
	};
}
