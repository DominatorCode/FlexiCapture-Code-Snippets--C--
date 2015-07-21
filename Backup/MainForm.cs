// © ABBYY. 2012.
// SAMPLES code is property of ABBYY, exclusive rights are reserved. 
// DEVELOPER is allowed to incorporate SAMPLES into his own APPLICATION and modify it 
// under the terms of License Agreement between ABBYY and DEVELOPER.

// Product: ABBYY FlexiCapture Engine 10

using System;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using System.Reflection;
using System.Reflection.Emit;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;

namespace Sample
{
	public class MainForm : Form, IProcessCallback
	{
		private MainMenu mainMenu;
		private MenuItem menuItemFile;
		private MenuItem menuItemExit;
		private MenuItem menuItemSnippet;
		private MenuItem menuItemRun;
		private MenuItem menuItemRunAll;
		private MenuItem menuItemRunRepeatedly;
		private MenuItem menuItemRunAllRepeatedly;
		private MenuItem menuItemRunTests;
		private MenuItem menuItemSplitter;
		private MenuItem menuItemViewSource;
		private MenuItem menuItemOptions;
		private MenuItem menuItemLoadNakedMode;
		private MenuItem menuItemLoadInprocServerMode;
		private MenuItem menuItemLoadWorkprocessMode;
		private MenuItem menuItemSplitter2;
		private MenuItem menuItemReuseLoadedEngine;
		private MenuItem menuItemMeasureTime;
		private Button buttonRun;
		private TreeView treeView;
		private ImageList imageList;
		private TreeView treeViewErrorsLog;
		private StatusBar statusBar;
		private ProgressBar progressBar;
		private PictureBox pictureBox;
		private System.Windows.Forms.Label labelHint1;
		private System.Windows.Forms.Label labelHint2;
		private System.Windows.Forms.Label labelHint3;
		private Button buttonViewSource;
		private System.Windows.Forms.Label labelLogHeader;
		private System.ComponentModel.IContainer components;
		
		public MainForm()
		{
			InitializeComponent();
		}

		private void MainForm_Load(object sender, System.EventArgs e)
		{
			// Will be updating state of controls in the application idle cycle
			idleHandler = new EventHandler( IdleUIUpdate );
			Application.Idle += idleHandler;

			// All tracing events will be processed by this form
			TracingSupport.SetCallback( this );	
			// The progress bar will be used to show progress when applicable 
			progressBar = new ProgressBar();
			progressBar.Dock = DockStyle.Right;
			statusBar.Controls.Add( progressBar );
			
			// Fill the tree with sets of snippets
			addSnippetsNode( "Common processing scenarios", FlexiCaptureEngineSnippets.GetSnippets( typeof( UsingFlexiCaptureTechnologyAPI ) ), true );
			addSnippetsNode( "Processing with a preconfigured FlexiCapture project", FlexiCaptureEngineSnippets.GetSnippets( typeof( WorkingWithFlexiCaptureProject ) ), true );
			addSnippetsNode( "Advanced techniques", FlexiCaptureEngineSnippets.GetSnippets( typeof( AdvancedTechniques ) ), true );
			addSnippetsNode( "Frequently asked questions", FlexiCaptureEngineSnippets.GetSnippets( typeof( FrequentlyAskedQuestions ) ), true );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MainForm));
			this.buttonRun = new System.Windows.Forms.Button();
			this.treeView = new System.Windows.Forms.TreeView();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.pictureBox = new System.Windows.Forms.PictureBox();
			this.treeViewErrorsLog = new System.Windows.Forms.TreeView();
			this.statusBar = new System.Windows.Forms.StatusBar();
			this.mainMenu = new System.Windows.Forms.MainMenu();
			this.menuItemFile = new System.Windows.Forms.MenuItem();
			this.menuItemExit = new System.Windows.Forms.MenuItem();
			this.menuItemSnippet = new System.Windows.Forms.MenuItem();
			this.menuItemRun = new System.Windows.Forms.MenuItem();
			this.menuItemRunRepeatedly = new System.Windows.Forms.MenuItem();
			this.menuItemRunAll = new System.Windows.Forms.MenuItem();
			this.menuItemRunAllRepeatedly = new System.Windows.Forms.MenuItem();
			this.menuItemRunTests = new System.Windows.Forms.MenuItem();
			this.menuItemSplitter = new System.Windows.Forms.MenuItem();
			this.menuItemViewSource = new System.Windows.Forms.MenuItem();
			this.menuItemOptions = new System.Windows.Forms.MenuItem();
			this.menuItemLoadNakedMode = new System.Windows.Forms.MenuItem();
			this.menuItemLoadInprocServerMode = new System.Windows.Forms.MenuItem();
			this.menuItemLoadWorkprocessMode = new System.Windows.Forms.MenuItem();
			this.menuItemSplitter2 = new System.Windows.Forms.MenuItem();
			this.menuItemReuseLoadedEngine = new System.Windows.Forms.MenuItem();
			this.menuItemMeasureTime = new System.Windows.Forms.MenuItem();
			this.buttonViewSource = new System.Windows.Forms.Button();
			this.labelHint1 = new System.Windows.Forms.Label();
			this.labelHint2 = new System.Windows.Forms.Label();
			this.labelHint3 = new System.Windows.Forms.Label();
			this.labelLogHeader = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// buttonRun
			// 
			this.buttonRun.Location = new System.Drawing.Point(176, 8);
			this.buttonRun.Name = "buttonRun";
			this.buttonRun.TabIndex = 0;
			this.buttonRun.Text = "Run";
			this.buttonRun.Click += new System.EventHandler(this.command_Run);
			// 
			// treeView
			// 
			this.treeView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.treeView.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(204)));
			this.treeView.HideSelection = false;
			this.treeView.ImageList = this.imageList;
			this.treeView.Indent = 23;
			this.treeView.Location = new System.Drawing.Point(176, 40);
			this.treeView.Name = "treeView";
			this.treeView.ShowRootLines = false;
			this.treeView.Size = new System.Drawing.Size(566, 753);
			this.treeView.TabIndex = 1;
			// 
			// imageList
			// 
			this.imageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
			this.imageList.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// pictureBox
			// 
			this.pictureBox.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(239)), ((System.Byte)(69)), ((System.Byte)(83)));
			this.pictureBox.Dock = System.Windows.Forms.DockStyle.Left;
			this.pictureBox.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox.Image")));
			this.pictureBox.Location = new System.Drawing.Point(0, 0);
			this.pictureBox.Name = "pictureBox";
			this.pictureBox.Size = new System.Drawing.Size(168, 793);
			this.pictureBox.TabIndex = 2;
			this.pictureBox.TabStop = false;
			// 
			// treeViewErrorsLog
			// 
			this.treeViewErrorsLog.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.treeViewErrorsLog.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(204)));
			this.treeViewErrorsLog.ImageIndex = -1;
			this.treeViewErrorsLog.Location = new System.Drawing.Point(176, 745);
			this.treeViewErrorsLog.Name = "treeViewErrorsLog";
			this.treeViewErrorsLog.SelectedImageIndex = -1;
			this.treeViewErrorsLog.ShowPlusMinus = false;
			this.treeViewErrorsLog.Size = new System.Drawing.Size(566, 48);
			this.treeViewErrorsLog.TabIndex = 1;
			this.treeViewErrorsLog.Visible = false;
			// 
			// statusBar
			// 
			this.statusBar.Location = new System.Drawing.Point(0, 793);
			this.statusBar.Name = "statusBar";
			this.statusBar.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.statusBar.Size = new System.Drawing.Size(750, 22);
			this.statusBar.TabIndex = 3;
			this.statusBar.Text = "Ready";
			// 
			// mainMenu
			// 
			this.mainMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					 this.menuItemFile,
																					 this.menuItemSnippet,
																					 this.menuItemOptions});
			// 
			// menuItemFile
			// 
			this.menuItemFile.Index = 0;
			this.menuItemFile.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						 this.menuItemExit});
			this.menuItemFile.Text = "&File";
			// 
			// menuItemExit
			// 
			this.menuItemExit.Index = 0;
			this.menuItemExit.Text = "E&xit";
			this.menuItemExit.Click += new System.EventHandler(this.MainForm_Close);
			// 
			// menuItemSnippet
			// 
			this.menuItemSnippet.Index = 1;
			this.menuItemSnippet.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																							this.menuItemRun,
																							this.menuItemRunRepeatedly,
																							this.menuItemRunAll,
																							this.menuItemRunAllRepeatedly,
																							this.menuItemRunTests,
																							this.menuItemSplitter,
																							this.menuItemViewSource});
			this.menuItemSnippet.Text = "&Snippet";
			// 
			// menuItemRun
			// 
			this.menuItemRun.Index = 0;
			this.menuItemRun.Shortcut = System.Windows.Forms.Shortcut.F5;
			this.menuItemRun.Text = "&Run selected";
			this.menuItemRun.Click += new System.EventHandler(this.command_Run);
			// 
			// menuItemRunRepeatedly
			// 
			this.menuItemRunRepeatedly.Index = 1;
			this.menuItemRunRepeatedly.Shortcut = System.Windows.Forms.Shortcut.CtrlP;
			this.menuItemRunRepeatedly.Text = "Run selected re&peatedly...";
			this.menuItemRunRepeatedly.Click += new System.EventHandler(this.command_RunRepeatedly);
			// 
			// menuItemRunAll
			// 
			this.menuItemRunAll.Index = 2;
			this.menuItemRunAll.Shortcut = System.Windows.Forms.Shortcut.CtrlL;
			this.menuItemRunAll.Text = "Run &all";
			this.menuItemRunAll.Click += new System.EventHandler(this.command_RunAll);
			// 
			// menuItemRunAllRepeatedly
			// 
			this.menuItemRunAllRepeatedly.Index = 3;
			this.menuItemRunAllRepeatedly.Text = "Run all repeatedly...";
			this.menuItemRunAllRepeatedly.Click += new System.EventHandler(this.command_RunAllRepeatedly);
			// 
			// menuItemRunTests
			// 
			this.menuItemRunTests.Index = 4;
			this.menuItemRunTests.Text = "Run long tests...";
			this.menuItemRunTests.Click += new System.EventHandler(this.command_RunTests);
			// 
			// menuItemSplitter
			// 
			this.menuItemSplitter.Index = 5;
			this.menuItemSplitter.Text = "-";
			// 
			// menuItemViewSource
			// 
			this.menuItemViewSource.Index = 6;
			this.menuItemViewSource.Shortcut = System.Windows.Forms.Shortcut.CtrlS;
			this.menuItemViewSource.Text = "View &source";
			this.menuItemViewSource.Click += new System.EventHandler(this.command_ViewSource);
			// 
			// menuItemOptions
			// 
			this.menuItemOptions.Index = 2;
			this.menuItemOptions.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																							this.menuItemLoadNakedMode,
																							this.menuItemLoadInprocServerMode,
																							this.menuItemLoadWorkprocessMode,
																							this.menuItemSplitter2,
																							this.menuItemReuseLoadedEngine,
																							this.menuItemMeasureTime});
			this.menuItemOptions.Text = "&Options";
			// 
			// menuItemLoadNakedMode
			// 
			this.menuItemLoadNakedMode.Index = 0;
			this.menuItemLoadNakedMode.RadioCheck = true;
			this.menuItemLoadNakedMode.Text = "Load Engine &directly and use naked interfaces";
			this.menuItemLoadNakedMode.Click += new System.EventHandler(this.menuItemLoadNakedMode_Click);
			// 
			// menuItemLoadInprocServerMode
			// 
			this.menuItemLoadInprocServerMode.Index = 1;
			this.menuItemLoadInprocServerMode.RadioCheck = true;
			this.menuItemLoadInprocServerMode.Text = "Load Engine as &in-process server";
			this.menuItemLoadInprocServerMode.Click += new System.EventHandler(this.menuItemLoadInprocServerMode_Click);
			// 
			// menuItemLoadWorkprocessMode
			// 
			this.menuItemLoadWorkprocessMode.Index = 2;
			this.menuItemLoadWorkprocessMode.RadioCheck = true;
			this.menuItemLoadWorkprocessMode.Text = "Load Engine as &out-of-process server in a worker process";
			this.menuItemLoadWorkprocessMode.Click += new System.EventHandler(this.menuItemLoadWorkprocessMode_Click);
			// 
			// menuItemSplitter2
			// 
			this.menuItemSplitter2.Index = 3;
			this.menuItemSplitter2.Text = "-";
			// 
			// menuItemReuseLoadedEngine
			// 
			this.menuItemReuseLoadedEngine.Index = 4;
			this.menuItemReuseLoadedEngine.RadioCheck = true;
			this.menuItemReuseLoadedEngine.Text = "&Reuse loaded Engine (do not reload for every snippet)";
			this.menuItemReuseLoadedEngine.Click += new System.EventHandler(this.menuItemReuseLoadedEngine_Click);
			// 
			// menuItemMeasureTime
			// 
			this.menuItemMeasureTime.Index = 5;
			this.menuItemMeasureTime.RadioCheck = true;
			this.menuItemMeasureTime.Text = "&Measure snippets running time (profiling)";
			this.menuItemMeasureTime.Click += new System.EventHandler(this.menuItemMeasureTime_Click);
			// 
			// buttonViewSource
			// 
			this.buttonViewSource.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.buttonViewSource.Location = new System.Drawing.Point(630, 8);
			this.buttonViewSource.Name = "buttonViewSource";
			this.buttonViewSource.Size = new System.Drawing.Size(112, 23);
			this.buttonViewSource.TabIndex = 4;
			this.buttonViewSource.Text = "View Source";
			this.buttonViewSource.Click += new System.EventHandler(this.command_ViewSource);
			// 
			// labelHint1
			// 
			this.labelHint1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(239)), ((System.Byte)(69)), ((System.Byte)(83)));
			this.labelHint1.Font = new System.Drawing.Font("Arial", 10F);
			this.labelHint1.ForeColor = System.Drawing.Color.White;
			this.labelHint1.Location = new System.Drawing.Point(8, 288);
			this.labelHint1.Name = "labelHint1";
			this.labelHint1.Size = new System.Drawing.Size(152, 112);
			this.labelHint1.TabIndex = 5;
			this.labelHint1.Text = "On the right you can see a list of short samples that show how ABBYY FlexiCapture" +
				" Engine should be used in different scenarios.";
			// 
			// labelHint2
			// 
			this.labelHint2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(239)), ((System.Byte)(69)), ((System.Byte)(83)));
			this.labelHint2.Font = new System.Drawing.Font("Arial", 10F);
			this.labelHint2.ForeColor = System.Drawing.Color.White;
			this.labelHint2.Location = new System.Drawing.Point(8, 400);
			this.labelHint2.Name = "labelHint2";
			this.labelHint2.Size = new System.Drawing.Size(152, 136);
			this.labelHint2.TabIndex = 5;
			this.labelHint2.Text = "Select a sample and run it to see a step-by-step execution log. You can view sour" +
				"ce code for the selected sample or any of the steps in the log.";
			// 
			// labelHint3
			// 
			this.labelHint3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(239)), ((System.Byte)(69)), ((System.Byte)(83)));
			this.labelHint3.Font = new System.Drawing.Font("Arial", 10F);
			this.labelHint3.ForeColor = System.Drawing.Color.White;
			this.labelHint3.ImageAlign = System.Drawing.ContentAlignment.TopRight;
			this.labelHint3.Location = new System.Drawing.Point(8, 528);
			this.labelHint3.Name = "labelHint3";
			this.labelHint3.Size = new System.Drawing.Size(152, 160);
			this.labelHint3.TabIndex = 5;
			this.labelHint3.Text = "Write a new sample or modify an existing one. Check how robust the Engine is in y" +
				"our scenario by setting it to run repeatedly. See \"How to test your scenario\" sa" +
				"mple for details.";
			// 
			// labelLogHeader
			// 
			this.labelLogHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.labelLogHeader.BackColor = System.Drawing.SystemColors.InactiveCaption;
			this.labelLogHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(204)));
			this.labelLogHeader.ForeColor = System.Drawing.SystemColors.InactiveCaptionText;
			this.labelLogHeader.Location = new System.Drawing.Point(176, 729);
			this.labelLogHeader.Name = "labelLogHeader";
			this.labelLogHeader.Size = new System.Drawing.Size(566, 16);
			this.labelLogHeader.TabIndex = 6;
			this.labelLogHeader.Text = "Execution Errors and Hints";
			this.labelLogHeader.Visible = false;
			// 
			// MainForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(750, 815);
			this.Controls.Add(this.treeViewErrorsLog);
			this.Controls.Add(this.labelLogHeader);
			this.Controls.Add(this.labelHint3);
			this.Controls.Add(this.labelHint2);
			this.Controls.Add(this.labelHint1);
			this.Controls.Add(this.buttonViewSource);
			this.Controls.Add(this.pictureBox);
			this.Controls.Add(this.treeView);
			this.Controls.Add(this.buttonRun);
			this.Controls.Add(this.statusBar);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Menu = this.mainMenu;
			this.MinimumSize = new System.Drawing.Size(700, 800);
			this.Name = "MainForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "FlexiCapture Engine Code Snippets";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.MainForm_Closing);
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.ResumeLayout(false);

		}
		#endregion

		// REQUIREMENT: The thread calling FCEngine must be a Single-Threaded Appartment (STA) thread
		[STAThread]
		static void Main() 
		{
			Application.Run( new MainForm() );
		}

		private void command_Run( object sender, System.EventArgs e )
		{
			// Run the selected snippet
			try {
				snippetIsRunning = true;
				Cursor = Cursors.WaitCursor;
				UIUpdate();

				TreeNode node = treeView.SelectedNode;
				if( node != null ) {
					Snippet snippet = node.Tag as Snippet;
					if( snippet != null ) {
						traceReset( node );
						DateTime started = DateTime.Now;
						FlexiCaptureEngineSnippets.Run( snippet, 1, true, engineLoadingMode, reuseLoadedEngine );		
						if( measureSnippetsRunningTime ) {
							int elapsedTime = (int)(DateTime.Now - started).TotalMilliseconds;
							tracePerformance( "Completed in: " + elapsedTime.ToString() + " ms" );
						}
					}
				}
				treeView.SelectedNode = node;
			} finally {
				FlexiCaptureEngineSnippets.UnloadEngine();
				Cursor = Cursors.Default;
				snippetIsRunning = false;
			}
		}

		#region ADDITIONAL COMMANDS

		private void command_RunAll(object sender, System.EventArgs e)
		{
			// Run all snippets
			try {
				snippetIsRunning = true;
				Cursor = Cursors.WaitCursor;
				UIUpdate();
				
				DateTime started0 = DateTime.Now;
				
				foreach( TreeNode collection in treeView.Nodes ) {
					foreach( TreeNode node in collection.Nodes ) {
						Snippet snippet = node.Tag as Snippet;
						if( snippet != null ) {
							traceReset( node );
							DateTime started = DateTime.Now;
							FlexiCaptureEngineSnippets.Run( snippet, 1, false, engineLoadingMode, reuseLoadedEngine );		
							if( measureSnippetsRunningTime ) {
								int elapsedTime = (int)(DateTime.Now - started).TotalMilliseconds;
								tracePerformance( "Completed in: " + elapsedTime.ToString() + " ms" );
							}
						}
					}
				}

				if( measureSnippetsRunningTime ) {
					int elapsedTime = (int)(DateTime.Now - started0).TotalMilliseconds;
					addToErrorsLog( "Completed in: " + elapsedTime.ToString() + " ms" );
				}

				treeView.SelectedNode = null;
			} finally {
				FlexiCaptureEngineSnippets.UnloadEngine();
				Cursor = Cursors.Default;
				snippetIsRunning = false;
			}		
		}

		private void command_RunRepeatedly( object sender, System.EventArgs e )
		{
			// Run the selected snippet repeatedly
			
			InputDialog inputDialog = new InputDialog( "Run repeatedly", "Repeat count:", "1",
				new InputDialog.InputValidationMethod( checkRepeatCount ) );

			if( inputDialog.ShowDialog() == DialogResult.OK ) {
				int count = int.Parse( inputDialog.InputText );
				try {
					snippetIsRunning = true;
					Cursor = Cursors.WaitCursor;
					UIUpdate();

					TreeNode node = treeView.SelectedNode;	
					
					if( node != null ) {
						Snippet snippet = node.Tag as Snippet;
						if( snippet != null ) {
							traceReset( node );
							DateTime started = DateTime.Now;
							FlexiCaptureEngineSnippets.Run( snippet, count, false, engineLoadingMode, reuseLoadedEngine );
							if( measureSnippetsRunningTime ) {
								int elapsedTime = (int)(DateTime.Now - started).TotalMilliseconds;
								tracePerformance( "Completed in: " + elapsedTime.ToString() + " ms" );
							}
						}
					}

					treeView.SelectedNode = node;
				} finally {
					FlexiCaptureEngineSnippets.UnloadEngine();
					Cursor = Cursors.Default;
					snippetIsRunning = false;
				}
			}
		}

		private void command_RunAllRepeatedly(object sender, System.EventArgs e)
		{
			// Run all snippets repeatedly

			InputDialog inputDialog = new InputDialog( "Run all repeatedly", "Repeat count:", "1",
				new InputDialog.InputValidationMethod( checkRepeatCount ) );

			if( inputDialog.ShowDialog() == DialogResult.OK ) {
				int count = int.Parse( inputDialog.InputText );
				runAllRepeatedly( count );
			}
		}
		
		private void command_RunTests(object sender, System.EventArgs e) 
		{
			// Burnin tests: Run all snippers in all engine loading modes repeatedly
			
			InputDialog inputDialog = new InputDialog( "Run long tests", "Repeat count:", "1",
				new InputDialog.InputValidationMethod( checkRepeatCount ) );

			if( inputDialog.ShowDialog() == DialogResult.OK ) {
				int count = int.Parse( inputDialog.InputText );
				EngineLoadingMode _engineLoadingMode = engineLoadingMode;
				bool _reuseLoadedEngine = reuseLoadedEngine;
				try {
					reuseLoadedEngine = true;
					addToErrorsLog( "Load Engine directly and use naked interfaces" );
					engineLoadingMode = EngineLoadingMode.LoadDirectlyUseNakedInterfaces;
					runAllRepeatedly( count );
					addToErrorsLog( "Load Engine as in-process server" );
					engineLoadingMode = EngineLoadingMode.LoadAsInprocServer;
					runAllRepeatedly( count );
					addToErrorsLog( "Load Engine in a worker process" );
					engineLoadingMode = EngineLoadingMode.LoadAsWorkprocess;
					runAllRepeatedly( count );
				} finally {
					engineLoadingMode = _engineLoadingMode;
					reuseLoadedEngine = _reuseLoadedEngine;
				}
			}
		}

		private void runAllRepeatedly( int repeatCount )
		{
			try {
				snippetIsRunning = true;
				Cursor = Cursors.WaitCursor;
				UIUpdate();
				
				DateTime started0 = DateTime.Now;
				
				foreach( TreeNode collection in treeView.Nodes ) {
					foreach( TreeNode node in collection.Nodes ) {
						Snippet snippet = node.Tag as Snippet;
						if( snippet != null ) {
							traceReset( node );
							DateTime started = DateTime.Now;
							FlexiCaptureEngineSnippets.Run( snippet, repeatCount, false, engineLoadingMode, reuseLoadedEngine );		
							if( measureSnippetsRunningTime ) {
								int elapsedTime = (int)(DateTime.Now - started).TotalMilliseconds;
								tracePerformance( "Completed in: " + elapsedTime.ToString() + " ms" );
							}
						}
					}
				}

				int elapsedTime0 = (int)(DateTime.Now - started0).TotalMilliseconds;
				addToErrorsLog( "Completed in: " + elapsedTime0.ToString() + " ms" );

				treeView.SelectedNode = null;
			} finally {
				FlexiCaptureEngineSnippets.UnloadEngine();
				Cursor = Cursors.Default;
				snippetIsRunning = false;
			}
		}

		private void command_ViewSource( object sender, System.EventArgs e ) 
		{
			// Open source file for the selected node in the tree

			try {
				TreeNode node = treeView.SelectedNode;
				if( node != null ) {
					if( node.Tag is Snippet ) {
						// Find entry point for the snippet
						Snippet snippet = node.Tag as Snippet;
						string fileName;
						int lineNumber;
						FlexiCaptureEngineSnippets.FindEntryPoint( snippet, out fileName, out lineNumber );
						if( lineNumber > 0 ) {
							openSourceFile( makeSourcePathExeRelative( fileName ), lineNumber );
						}
					} else if( node.Tag is TraceData ) {
						// Find the trace source in the source code
						TraceData traceData = (TraceData)node.Tag;
						openSourceFile( traceData.FileName, traceData.LineNumber );
					}
				}	
			} catch( Exception ex ) {
				MessageBox.Show( ex.Message );
			}
		}

		private void menuItemLoadNakedMode_Click(object sender, System.EventArgs e)
		{
			engineLoadingMode = EngineLoadingMode.LoadDirectlyUseNakedInterfaces;
		}

		private void menuItemLoadInprocServerMode_Click(object sender, System.EventArgs e) 
		{
			engineLoadingMode = EngineLoadingMode.LoadAsInprocServer;
		}

		private void menuItemLoadWorkprocessMode_Click(object sender, System.EventArgs e) 
		{
			engineLoadingMode = EngineLoadingMode.LoadAsWorkprocess;
		}

		private void menuItemReuseLoadedEngine_Click(object sender, System.EventArgs e) 
		{
			reuseLoadedEngine = !reuseLoadedEngine;
		}

		private void menuItemMeasureTime_Click(object sender, System.EventArgs e)
		{
			measureSnippetsRunningTime = !measureSnippetsRunningTime;
		}

		#endregion

		#region IMPLEMENTATION

		bool snippetIsRunning = false;
		EventHandler idleHandler = null;
		EngineLoadingMode engineLoadingMode = EngineLoadingMode.LoadDirectlyUseNakedInterfaces;
		bool reuseLoadedEngine = true;
		bool measureSnippetsRunningTime = false;

		protected override void Dispose( bool disposing )
		{
			// Perform required cleanup before closing
			if( disposing ) {
				TracingSupport.SetCallback( null );
				Application.Idle -= idleHandler;
			}
			base.Dispose( disposing );
		}

		void MainForm_Close( object sender, System.EventArgs e ) 
		{
			Close();
		}
		
		void MainForm_Closing( object sender, System.ComponentModel.CancelEventArgs e ) 
		{
			// Prevent the form from closing if a snippet is running
			if( snippetIsRunning ) {
				e.Cancel = true;
			}
		}

		void IdleUIUpdate( object sender, EventArgs e )
		{
			UIUpdate();
		}

		void UIUpdate()
		{
			// Update controls in the application idle cycle
			if( !snippetIsRunning ) {
				TreeNode node = treeView.SelectedNode;
				if( node != null && node.Tag is Snippet ) {
					buttonRun.Enabled = true;
					menuItemRun.Enabled = true;
					menuItemRunRepeatedly.Enabled = true;
					buttonViewSource.Enabled = true;
					menuItemViewSource.Enabled = true;
					statusBar.Text = "Press 'Run' (F5) button to run the selected scenario or 'View Source' to see the source code.";
				} else {
					buttonRun.Enabled = false;
					menuItemRun.Enabled = false;
					menuItemRunRepeatedly.Enabled = false;
					menuItemRunRepeatedly.Enabled = false;
					if( node != null && node.Tag is TraceData ) {
						TraceData traceData = (TraceData)node.Tag;
						statusBar.Text = traceData.FileName + ", " + traceData.LineNumber;
						buttonViewSource.Enabled = true;
						menuItemViewSource.Enabled = true;
					} else {
						statusBar.Text = "Select a scenario.";
						buttonViewSource.Enabled = false;
						menuItemViewSource.Enabled = false;
					}
				}
				menuItemRunAll.Enabled = true;
				menuItemRunAllRepeatedly.Enabled = true;
				menuItemRunTests.Enabled = true;
				progressBar.Visible = false;
				
				menuItemLoadNakedMode.Enabled = true;
				menuItemLoadNakedMode.Checked = ( engineLoadingMode == EngineLoadingMode.LoadDirectlyUseNakedInterfaces );
				menuItemLoadInprocServerMode.Enabled = true;
				menuItemLoadInprocServerMode.Checked = ( engineLoadingMode == EngineLoadingMode.LoadAsInprocServer );
				menuItemLoadWorkprocessMode.Enabled = true;
				menuItemLoadWorkprocessMode.Checked = ( engineLoadingMode == EngineLoadingMode.LoadAsWorkprocess );

				menuItemReuseLoadedEngine.Enabled = true;
				menuItemReuseLoadedEngine.Checked = reuseLoadedEngine;
				menuItemMeasureTime.Enabled = true;
				menuItemMeasureTime.Checked = measureSnippetsRunningTime;
			} else {
				buttonRun.Enabled = false;
				menuItemRun.Enabled = false;
				menuItemRunRepeatedly.Enabled = false;
				menuItemRunAll.Enabled = false;
				menuItemRunAllRepeatedly.Enabled = false;
				menuItemRunTests.Enabled = false;
				menuItemLoadNakedMode.Enabled = false;
				menuItemLoadInprocServerMode.Enabled = false;
				menuItemLoadWorkprocessMode.Enabled = false;
				menuItemReuseLoadedEngine.Enabled = false;
				buttonViewSource.Enabled = false;
				menuItemViewSource.Enabled = false;
				menuItemMeasureTime.Enabled = false;

				switch( engineLoadingMode ) {
					case EngineLoadingMode.LoadDirectlyUseNakedInterfaces: 
						statusBar.Text = "Running...";
						break;
					case EngineLoadingMode.LoadAsInprocServer:
						statusBar.Text = "Running (inproc server mode)...";
						break;
					case EngineLoadingMode.LoadAsWorkprocess:
						statusBar.Text = "Running (workprocess server mode)...";
						break;
				}
			}
		}

		void addSnippetsNode( string text, Snippet[] snippets, bool expand )
		{
			// Create a top level node in the tree view and add the snippets as child nodes
			TreeNode folderNode = treeView.Nodes.Add( text );
			foreach( Snippet snippet in snippets ) {
				string methodName = snippet.Method.Name;
				TreeNode node = folderNode.Nodes.Add( methodName.Replace( "LPAREN", "(" ).Replace( "RPAREN", ")" ).Replace( "_", " " ) );
				node.ImageIndex = node.SelectedImageIndex = 1;
				node.Tag = snippet;
			}
			if( expand ) {
				folderNode.Expand();
			}
		}

		bool checkRepeatCount( string text, out string errorMessage )
		{
			errorMessage = "";
			try {
				int Value = int.Parse( text );
				if( Value < 1 ) { 
					errorMessage = "Must be one or greater";
					return false;
				}
				return true;
			} catch( Exception ) {
				errorMessage = "Not a valid numeric value";
				return false;
			}
		}

		void tracePerformance( string message )
		{
			( (IProcessCallback)this ).Trace( message, "", 0, true  );
		}

		private void addToErrorsLog( string message )
		{
			treeView.Height = labelLogHeader.Top - treeView.Top;
			labelLogHeader.Show();
			treeViewErrorsLog.Show();
			treeViewErrorsLog.Nodes.Add( message );
		}
		
		#region	/// IProcessCallback //////////////////////////////////////////////
		
		// Icon indices in the imageList
		const int TraceOkIconIndex = 2;
		const int TraceErrorIconIndex = 3;

		class TraceData
		{
			public string FileName;
			public int LineNumber;
			public int Level;

			public TraceData( string fileName, int lineNumber, int level ) 
			{ 
				FileName = fileName;
				LineNumber = lineNumber;
				Level = level; 
			}
		};

		TreeNode currentNode;

		void IProcessCallback.Trace( string message, string fileName, int lineNumber, bool performanceIndicator )
		{
			if( measureSnippetsRunningTime && !performanceIndicator ) {
				return;
			}
			// Show trace messages in a TreeView
			TreeNode node = currentNode.Nodes.Add( message );
			node.ImageIndex = node.SelectedImageIndex = TraceOkIconIndex;
			currentNode.Expand();
			
			node.Tag = new TraceData( makeSourcePathExeRelative( fileName ), lineNumber, -1 );

			node.EnsureVisible();
			node.TreeView.SelectedNode = node;
			
			Refresh();
			Application.DoEvents();
		}

		void IProcessCallback.TraceBegin( string message, string fileName, int lineNumber, int level, bool performanceIndicator )
		{
			if( measureSnippetsRunningTime && !performanceIndicator ) {
				return;
			}
			
			// Begin trace block
			currentNode = currentNode.Nodes.Add( message );
			currentNode.Tag = new TraceData( makeSourcePathExeRelative( fileName ), lineNumber, level );
			currentNode.ImageIndex = currentNode.SelectedImageIndex = TraceOkIconIndex;
			currentNode.Parent.Expand();
			
			progressBar.Value = 0;

			currentNode.EnsureVisible();
			currentNode.TreeView.SelectedNode = currentNode;

			Refresh();
			Application.DoEvents();
		}

		void IProcessCallback.TraceEnd( string message, int level, bool performanceIndicator )
		{
			if( measureSnippetsRunningTime && !performanceIndicator ) {
				return;
			}
			
			// End normal trace block
			traceEnd( message, level, false );

			Refresh();
			Application.DoEvents();
		}

		void IProcessCallback.TraceException( string message, string fileName, int lineNumber, int level )
		{
			// Show exception and unwind the "tracing stack"
			addToErrorsLog( message );
			statusBar.Text = message;

			TreeNode node = currentNode.Nodes.Add( message );
			node.ImageIndex = node.SelectedImageIndex = TraceErrorIconIndex;
			node.Tag = new TraceData( makeSourcePathExeRelative( fileName ), lineNumber, -1 );
			currentNode.Expand();
			
			// Ånd all trace blocks up to the level where the exception is caught
			traceEnd( "", level, true );

			Refresh();
			Application.DoEvents();
		}

		void IProcessCallback.TraceProgress( int percent )
		{
			if( measureSnippetsRunningTime ) {
				return;
			}
			
			// Show progress in the progress bar
			progressBar.Value = percent;
			progressBar.Visible = true;
			
			Refresh();
			Application.DoEvents();
		}

		void IProcessCallback.TraceDetail( string message )
		{
			if( measureSnippetsRunningTime ) {
				return;
			}
			
			// Show details of the current operation in the status bar
			statusBar.Text = message;
			
			Refresh();
			Application.DoEvents();
		}

		public void traceReset( TreeNode node ) 
		{
			// Reset tracing
			node.Nodes.Clear();
			currentNode = node;
			progressBar.Value = 0;
			progressBar.Visible = false;

			Refresh();
			Application.DoEvents();
		}

		public void traceEnd( string message, int level, bool isError )
		{
			while( currentNode.Tag is TraceData && ((TraceData)currentNode.Tag).Level >= level ) {
				currentNode.Text += message;
				if( isError ) {
					currentNode.ImageIndex = currentNode.SelectedImageIndex = TraceErrorIconIndex;
				} else {
					currentNode.Collapse();
				}
				currentNode = currentNode.Parent;
				
				if( !isError && currentNode.Tag is TraceData && ((TraceData)currentNode.Tag).Level == level ) { 
					break;
				}
			}

			progressBar.Value = 0;
			progressBar.Visible = false;
		}

		PictureBox graphPictureBox;
		int counter = 0;
		float prevValue;
        float scale;
		
		void IProcessCallback.TraceMemory( int memoryBytes ) 
		{
			if( graphPictureBox == null || graphPictureBox.Parent == null ) {
				Form form = new Form();
				form.Owner = this;
				form.Text = "Memory (Private Bytes)";
				form.FormBorderStyle = FormBorderStyle.FixedToolWindow;
				form.ShowInTaskbar = false;
				form.BackColor = Color.Black;
				graphPictureBox = new PictureBox();
				graphPictureBox.Size = form.ClientSize;
				graphPictureBox.Image = new Bitmap( graphPictureBox.ClientSize.Width, graphPictureBox.ClientSize.Height );

				using( Graphics g = Graphics.FromImage( graphPictureBox.Image ) ) {
					for( int i = 1; i < 10; i++ ) {
						int y = graphPictureBox.ClientSize.Height / 10 * i;
						g.DrawLine( Pens.DarkGray,  0,  y, graphPictureBox.ClientSize.Width,  y );
					}
				}

				form.Controls.Add( graphPictureBox );

				form.Show();
				form.SetDesktopLocation( this.DesktopLocation.X + Width, this.DesktopLocation.Y );

				counter = 0;
				prevValue = memoryBytes;
				scale = (float)( graphPictureBox.ClientSize.Height ) / memoryBytes / 2;
				return;
			}
			graphPictureBox.Parent.Show();
			
			const int xScale = 3;
			using( Graphics g = Graphics.FromImage( graphPictureBox.Image ) ) {
				if( prevValue * scale > graphPictureBox.ClientSize.Height ) {
					scale = (float)( graphPictureBox.ClientSize.Height ) / prevValue / 2;
					g.DrawLine( Pens.DarkGray,  xScale * counter,  0, xScale * counter,  graphPictureBox.ClientSize.Height );
				}
                PointF prevPoint = new PointF( xScale * counter, graphPictureBox.ClientSize.Height - prevValue * scale );
                prevValue = memoryBytes;
                PointF newPoint = new PointF( xScale * ( counter + 1 ), graphPictureBox.ClientSize.Height - prevValue * scale );
                g.DrawLine( Pens.Yellow, prevPoint, newPoint );
                prevPoint = newPoint;
            }
            graphPictureBox.Invalidate();
			
			counter++;
			if( ( counter + 1 ) * xScale > graphPictureBox.ClientSize.Width ) {
				counter = 0;
			}
		}

		#endregion

		#region VISUAL STUDIO INTEGRATION
		
		[DllImport( "user32.dll", CharSet=CharSet.Unicode ), PreserveSig]
		public static extern bool SetForegroundWindow( IntPtr hWnd );

		string makeSourcePathExeRelative( string filePath )
		{
			string _filePath = filePath;
			string exePath = Application.ExecutablePath;
			int pos = exePath.IndexOf( "\\bin\\" );
			if( pos != -1 ) {
				string projectDir = exePath.Substring( 0, pos + 1 );
				string fileName = Path.GetFileName( filePath );
				_filePath = Path.Combine( projectDir, fileName );
			}
			return _filePath;
		}

		void openSourceFile( string filePath, int lineNumber )
		{
			try {
				// If EnvDTE is not installed, an exception will be thrown during JIT-compilation 
				// of the next method. 
				_openSourceFile( filePath, lineNumber );
			} catch {
				MessageBox.Show( "Failed to connect to Microsoft Visual Studio EnvDTE.\n" +
					"Microsoft Visual Studio is either not installed or not compatible " +
					"with the compiled version of this sample.\n" +
					"Try recompiling the sample with your version of programming environment" );
			}
		}

		void _openSourceFile( string filePath, int lineNumber )
		{
			
			EnvDTE.DTE dte = getVisualStudioInstance();
			if( dte != null ) {
				dte.UserControl = true;
								
				if( System.IO.File.Exists( filePath ) ) {
					EnvDTE.Window window = dte.OpenFile( EnvDTE.Constants.vsViewKindTextView, filePath );
					window.Activate();
					
					EnvDTE.TextSelection textSelection = (EnvDTE.TextSelection)dte.ActiveDocument.Selection;
					textSelection.GotoLine( lineNumber - 1, true );
					
					dte.MainWindow.Activate();
					SetForegroundWindow( new IntPtr( dte.MainWindow.HWnd ) );
				} else {
					MessageBox.Show( "Couldn't find file " + filePath );
				}
				
			} else {
				MessageBox.Show( "Couldn't launch Microsoft Visual Studio" );
			}	
		}

		EnvDTE.DTE getVisualStudioInstance()
		{
			object obj = null;
			if( obj == null ) {
				// Visual Studio 2003
				try {
					obj = System.Runtime.InteropServices.Marshal.GetActiveObject( "VisualStudio.DTE.7.1" );
				} catch( Exception ) {
				}
			}
			if( obj == null ) {
				// Visual Studio 2005
				try {
					obj = System.Runtime.InteropServices.Marshal.GetActiveObject( "VisualStudio.DTE.8.0" );
				} catch( Exception ) {
				}
			}
			if( obj == null ) {
				// Visual Studio 2008
				try {
					obj = System.Runtime.InteropServices.Marshal.GetActiveObject( "VisualStudio.DTE.9.0" );
				} catch( Exception ) {
				}
			}
			if( obj == null ) {
				// Visual Studio 2010
				try {
					obj = System.Runtime.InteropServices.Marshal.GetActiveObject( "VisualStudio.DTE.10.0" );
				} catch( Exception ) {
				}
			}
			if( obj == null ) {
				System.Type type = System.Type.GetTypeFromProgID( "VisualStudio.DTE" );
				obj = System.Activator.CreateInstance( type );
			}
			return obj as EnvDTE.DTE;
		}

		#endregion

		#endregion
	}
}
