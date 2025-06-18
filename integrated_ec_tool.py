#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
統合ECツール - 商品登録から出品まで一括管理
"""

import sys
import os
import subprocess
import logging
from pathlib import Path
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTabWidget, QTextEdit, QLabel, QFileDialog,
    QMessageBox, QGroupBox, QGridLayout, QListWidget, QSplitter,
    QProgressBar, QStatusBar, QToolBar, QAction, QLineEdit, QComboBox,
    QInputDialog, QProgressDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, pyqtSlot
from PyQt5.QtGui import QIcon, QFont
import csv
import json
from datetime import datetime
import shutil
import time
import win32gui
import win32con
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, QTimer
import configparser

# 既存のproduct_appをインポート
try:
    from product_app import ProductApp
except ImportError:
    ProductApp = None

class IntegratedECTool(QMainWindow):
    """メインの統合ツールウィンドウ"""
    
    def __init__(self):
        super().__init__()
        self.master_tool_path = r"\\express5800\ITimpel\EXE\HAMST040.exe"  # マスタ管理ツール
        self.master_process = None  # HANBAIMENU.exeのプロセス
        self.master_hwnd = None     # HANBAIMENU.exeのウィンドウハンドル
        self.embed_timer = QTimer() # 埋め込み用タイマー
        self.embed_timer.timeout.connect(self.try_embed_master)
        self.resize_timer = QTimer()  # リサイズ用タイマー
        self.resize_timer.timeout.connect(self.resize_embedded_window)
        self.resize_timer.setSingleShot(True)
        self.last_resize_size = None  # 前回のサイズを記録
        self.embed_attempt_count = 0  # 埋め込み試行回数
        self.init_ui()
        self.setup_logging()
        
    def init_ui(self):
        """UI初期化"""
        self.setWindowTitle("統合EC管理ツール")
        
        # 画面サイズに応じてウィンドウサイズを調整
        screen = QApplication.primaryScreen().size()
        width = min(1600, int(screen.width() * 0.9))
        height = min(1000, int(screen.height() * 0.9))
        self.setGeometry(100, 100, width, height)
        
        # DPIスケールファクターを取得
        screen_obj = QApplication.primaryScreen()
        self.dpi_scale = screen_obj.physicalDotsPerInch() / 96.0
        
        # ステータスバー（先に作成）
        self.setStatusBar(QStatusBar())
        # self.statusBar().showMessage("起動中...")
        
        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # レイアウト
        layout = QVBoxLayout(main_widget)
        
        # ツールバー作成
        self.create_toolbar()
        
        # タブウィジェット
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)
        
        # 各タブを追加
        self.setup_workflow_tab()
        self.setup_master_tab()
        self.setup_product_tab()
        self.setup_upload_tab()
        self.setup_check_tab()
        
        # 初期タブをマスタタブに設定
        self.tabs.setCurrentIndex(1)  # マスタタブ
        
        # タブ切り替え時の処理を追加
        self.tabs.currentChanged.connect(self.on_tab_changed)
        
        # 準備完了
        # self.statusBar().showMessage("準備完了")
        
        # マスタツールを自動起動
        QTimer.singleShot(1000, self.auto_launch_master)
        
        # パスワード自動読み込み
        QTimer.singleShot(2000, self.load_ftp_password)
        
    def create_toolbar(self):
        """ツールバー作成"""
        toolbar = QToolBar()
        self.addToolBar(toolbar)
        
        # 設定アクション
        settings_action = QAction("設定", self)
        settings_action.triggered.connect(self.open_settings)
        toolbar.addAction(settings_action)
        
        # ヘルプアクション
        help_action = QAction("ヘルプ", self)
        help_action.triggered.connect(self.show_help)
        toolbar.addAction(help_action)
        
    def setup_workflow_tab(self):
        """ワークフロー管理タブ"""
        workflow_widget = QWidget()
        main_layout = QHBoxLayout(workflow_widget)
        
        # 左側パネル: ワークフロー制御
        left_panel = QWidget()
        panel_width = int(400 * self.dpi_scale)
        left_panel.setMaximumWidth(panel_width)
        left_layout = QVBoxLayout(left_panel)
        
        # ワークフロー手順
        workflow_group = QGroupBox("🔄 作業フロー")
        workflow_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        workflow_layout = QVBoxLayout(workflow_group)
        
        self.workflow_list = QListWidget()
        workflow_items = [
            "1. 📊 マスタ作成 (商品一覧)",
            "2. 📝 商品情報入力・CSV生成",
            "3. 🖼️ 商品画像準備",
            "4. 🔄 楽天: FTP アップロード",
            "5. 🌐 Yahoo: ストアクリエイター アップロード",
            "6. ✅ ページ確認・検証"
        ]
        for item in workflow_items:
            self.workflow_list.addItem(item)
        
        # リストアイテムのスタイル設定（DPI対応）
        font_size = max(10, int(12 * self.dpi_scale))
        padding = max(6, int(8 * self.dpi_scale))
        self.workflow_list.setStyleSheet(f"""
            QListWidget::item {{
                padding: {padding}px;
                border-bottom: 1px solid #e0e0e0;
                font-size: {font_size}px;
            }}
            QListWidget::item:selected {{
                background-color: #e3f2fd;
                color: #1976d2;
            }}
        """)
        workflow_layout.addWidget(self.workflow_list)
        
        left_layout.addWidget(workflow_group)
        
        # 実行制御
        control_group = QGroupBox("🚀 実行制御")
        control_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        control_layout = QVBoxLayout(control_group)
        
        self.auto_execute_btn = QPushButton("📤 アップロードワークフロー実行")
        btn_height = max(35, int(45 * self.dpi_scale))
        btn_font_size = max(11, int(13 * self.dpi_scale))
        border_radius = max(3, int(5 * self.dpi_scale))
        self.auto_execute_btn.setMinimumHeight(btn_height)
        self.auto_execute_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                border-radius: {border_radius}px;
                font-size: {btn_font_size}px;
            }}
            QPushButton:hover {{
                background-color: #45a049;
            }}
        """)
        self.auto_execute_btn.clicked.connect(self.auto_execute_workflow)
        control_layout.addWidget(self.auto_execute_btn)
        
        info_label = QLabel("💡 商品情報入力・CSV生成は「商品情報入力」タブで実行")
        info_label.setWordWrap(True)
        info_font_size = max(9, int(11 * self.dpi_scale))
        info_padding = max(8, int(10 * self.dpi_scale))
        info_border_radius = max(2, int(3 * self.dpi_scale))
        info_label.setStyleSheet(f"color: #666; font-size: {info_font_size}px; padding: {info_padding}px; background-color: #f5f5f5; border-radius: {info_border_radius}px;")
        control_layout.addWidget(info_label)
        
        left_layout.addWidget(control_group)
        left_layout.addStretch()
        
        main_layout.addWidget(left_panel)
        
        # 右側パネル: 進捗・ログ
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        # 進捗状況
        progress_group = QGroupBox("📈 進捗状況")
        progress_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        progress_layout = QVBoxLayout(progress_group)
        
        # 進捗バー
        progress_label = QLabel("全体進捗:")
        progress_label.setStyleSheet("font-weight: bold; margin-bottom: 5px;")
        progress_layout.addWidget(progress_label)
        
        self.progress_bar = QProgressBar()
        progress_height = max(20, int(25 * self.dpi_scale))
        progress_border = max(1, int(2 * self.dpi_scale))
        progress_radius = max(3, int(5 * self.dpi_scale))
        self.progress_bar.setMinimumHeight(progress_height)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: {progress_border}px solid #ddd;
                border-radius: {progress_radius}px;
                text-align: center;
                font-weight: bold;
            }}
            QProgressBar::chunk {{
                background-color: #4CAF50;
                border-radius: {max(2, int(3 * self.dpi_scale))}px;
            }}
        """)
        progress_layout.addWidget(self.progress_bar)
        
        right_layout.addWidget(progress_group)
        
        # ログ表示
        log_group = QGroupBox("📋 実行ログ")
        log_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText("ワークフロー実行ログがここに表示されます...")
        log_font_size = max(8, int(10 * self.dpi_scale))
        log_border_radius = max(2, int(3 * self.dpi_scale))
        self.log_text.setStyleSheet(f"""
            QTextEdit {{
                background-color: #fafafa;
                border: 1px solid #ddd;
                border-radius: {log_border_radius}px;
                font-family: 'Courier New', monospace;
                font-size: {log_font_size}px;
            }}
        """)
        log_layout.addWidget(self.log_text)
        
        right_layout.addWidget(log_group)
        
        main_layout.addWidget(right_panel, 1)
        
        self.tabs.addTab(workflow_widget, "ワークフロー")
        
    def setup_master_tab(self):
        """マスタ作成タブ"""
        master_widget = QWidget()
        layout = QVBoxLayout(master_widget)
        
        # メインレイアウト（横分割：埋め込みエリア + 小さなコントロール）
        main_horizontal_layout = QHBoxLayout()
        
        # HANBAIMENU.exe埋め込みエリア（メイン領域）
        self.master_embed_widget = QWidget()
        self.master_embed_widget.setStyleSheet("border: 2px solid gray; background-color: #f0f0f0;")
        embed_layout = QVBoxLayout(self.master_embed_widget)
        embed_layout.addWidget(QLabel("マスタ作成ツールを起動するとここに表示されます"))
        main_horizontal_layout.addWidget(self.master_embed_widget, 1)  # 最大領域を占有
        
        # 右側のコントロールパネル（全ての設定・説明・監視を集約）
        right_panel = QWidget()
        right_panel.setMaximumWidth(250)
        right_panel_layout = QVBoxLayout(right_panel)
        
        # 設定セクション
        settings_group = QGroupBox("設定")
        settings_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        settings_layout = QVBoxLayout(settings_group)
        
        # ツールパス表示（参照ボタンは削除 - 通常は固定パスのため不要）
        path_info = QLabel("ツールパス: HAMST040.exe")
        path_info.setStyleSheet("font-size: 10px; color: #666;")
        path_info.setWordWrap(True)
        settings_layout.addWidget(path_info)
        
        settings_group.setMaximumHeight(60)
        right_panel_layout.addWidget(settings_group)
        
        # マスタツール説明
        info_group = QGroupBox("について")
        info_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        info_layout = QVBoxLayout(info_group)
        
        info_text = QLabel("マスタ作成ツールで商品の基本データを管理。\n商品情報入力・CSV生成は「商品情報入力」タブで実行。")
        info_text.setWordWrap(True)
        info_text.setStyleSheet("font-size: 10px;")
        info_layout.addWidget(info_text)
        
        info_group.setMaximumHeight(80)
        right_panel_layout.addWidget(info_group)
        
        # 出力ファイル監視
        monitor_group = QGroupBox("ファイル監視")
        monitor_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        monitor_layout = QVBoxLayout(monitor_group)
        
        self.monitor_list = QListWidget()
        self.monitor_list.setMaximumHeight(100)
        self.monitor_list.setStyleSheet("font-size: 9px;")
        monitor_layout.addWidget(self.monitor_list)
        
        self.auto_import_check = QPushButton("自動取込")
        self.auto_import_check.setCheckable(True)
        self.auto_import_check.setMaximumHeight(25)
        self.auto_import_check.setStyleSheet("font-size: 9px;")
        monitor_layout.addWidget(self.auto_import_check)
        
        right_panel_layout.addWidget(monitor_group)
        
        # 手動リサイズコントロール
        resize_group = QGroupBox("リサイズ制御")
        resize_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        resize_layout = QVBoxLayout(resize_group)
        
        self.manual_resize_button = QPushButton("手動リサイズ")
        self.manual_resize_button.setMaximumHeight(25)
        self.manual_resize_button.setStyleSheet("font-size: 9px;")
        self.manual_resize_button.clicked.connect(self.manual_resize_master)
        resize_layout.addWidget(self.manual_resize_button)
        
        resize_group.setMaximumHeight(60)
        right_panel_layout.addWidget(resize_group)
        
        right_panel_layout.addStretch()  # 下に余白を追加
        main_horizontal_layout.addWidget(right_panel)
        
        layout.addLayout(main_horizontal_layout, 1)
        
        # リサイズイベントを監視
        self.master_embed_widget.resizeEvent = self.on_embed_widget_resize
        
        self.tabs.addTab(master_widget, "マスタ作成")
        
    def setup_product_tab(self):
        """商品情報入力タブ"""
        if ProductApp:
            # 既存のProductAppを組み込み
            self.product_app = ProductApp()
            self.tabs.addTab(self.product_app, "商品情報入力")
        else:
            # フォールバック
            product_widget = QWidget()
            layout = QVBoxLayout(product_widget)
            layout.addWidget(QLabel("商品情報入力機能（product_app.py）"))
            self.tabs.addTab(product_widget, "商品情報入力")
            
    def setup_image_tab(self):
        """画像管理タブ"""
        image_widget = QWidget()
        main_layout = QHBoxLayout(image_widget)
        
        # 左側パネル: フォルダ管理・操作
        left_panel = QWidget()
        panel_width = int(420 * self.dpi_scale)  # 幅を広げる
        left_panel.setMaximumWidth(panel_width)
        left_layout = QVBoxLayout(left_panel)
        
        # フォルダ設定
        folder_group = QGroupBox("📁 画像フォルダ設定")
        folder_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
                min-width: {int(250 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        folder_layout = QVBoxLayout(folder_group)
        
        # 現在のフォルダ表示
        current_folder_label = QLabel("現在のフォルダ:")
        current_folder_label.setStyleSheet("font-weight: bold; margin-bottom: 5px;")
        folder_layout.addWidget(current_folder_label)
        
        self.image_folder_label = QLabel("📂 未設定")
        self.image_folder_label.setWordWrap(True)
        self.image_folder_label.setStyleSheet("""
            QLabel {
                padding: 10px;
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 3px;
                color: #666;
                min-height: 20px;
            }
        """)
        folder_layout.addWidget(self.image_folder_label)
        
        # フォルダ選択ボタン
        browse_folder_btn = QPushButton("📂 フォルダを選択")
        btn_height = int(35 * self.dpi_scale)
        btn_font_size = int(12 * self.dpi_scale)
        border_radius = int(5 * self.dpi_scale)
        browse_folder_btn.setMinimumHeight(btn_height)
        browse_folder_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                border-radius: {border_radius}px;
                font-size: {btn_font_size}px;
            }}
            QPushButton:hover {{
                background-color: #1976D2;
            }}
        """)
        browse_folder_btn.clicked.connect(self.browse_image_folder)
        folder_layout.addWidget(browse_folder_btn)
        
        left_layout.addWidget(folder_group)
        
        # 操作パネル
        action_group = QGroupBox("🔧 操作")
        action_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        action_layout = QVBoxLayout(action_group)
        
        refresh_btn = QPushButton("🔄 画像リストを更新")
        refresh_btn.setMinimumHeight(btn_height)
        refresh_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #FF9800;
                color: white;
                font-weight: bold;
                border-radius: {border_radius}px;
                font-size: {btn_font_size}px;
            }}
            QPushButton:hover {{
                background-color: #F57C00;
            }}
        """)
        refresh_btn.clicked.connect(self.refresh_image_list)
        action_layout.addWidget(refresh_btn)
        
        # フォルダを開くボタン
        open_folder_btn = QPushButton("📁 フォルダを開く")
        open_folder_btn.setMinimumHeight(btn_height)
        open_folder_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                border-radius: {border_radius}px;
                font-size: {btn_font_size}px;
            }}
            QPushButton:hover {{
                background-color: #45a049;
            }}
        """)
        open_folder_btn.clicked.connect(self.open_current_image_folder)
        action_layout.addWidget(open_folder_btn)
        
        left_layout.addWidget(action_group)
        
        # 統計情報
        stats_group = QGroupBox("📊 統計情報")
        stats_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        stats_layout = QVBoxLayout(stats_group)
        
        self.image_count_label = QLabel("画像ファイル数: 0")
        self.image_count_label.setStyleSheet("font-weight: bold; color: #333;")
        stats_layout.addWidget(self.image_count_label)
        
        self.supported_formats_label = QLabel("対応形式: JPG, PNG, GIF, BMP")
        format_font_size = max(12, int(14 * self.dpi_scale))
        self.supported_formats_label.setStyleSheet(f"color: #666; font-size: {format_font_size}px;")
        stats_layout.addWidget(self.supported_formats_label)
        
        left_layout.addWidget(stats_group)
        left_layout.addStretch()
        
        main_layout.addWidget(left_panel)
        
        # 右側パネル: 画像リスト表示
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        # 画像一覧
        image_list_group = QGroupBox("🖼️ 画像ファイル一覧")
        image_list_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        image_list_layout = QVBoxLayout(image_list_group)
        
        self.image_list = QListWidget()
        list_font_size = int(11 * self.dpi_scale)
        list_padding = int(8 * self.dpi_scale)
        list_border_radius = int(3 * self.dpi_scale)
        self.image_list.setStyleSheet(f"""
            QListWidget {{
                background-color: #fafafa;
                border: 1px solid #ddd;
                border-radius: {list_border_radius}px;
                font-size: {list_font_size}px;
            }}
            QListWidget::item {{
                padding: {list_padding}px;
                border-bottom: 1px solid #eee;
            }}
            QListWidget::item:selected {{
                background-color: #e3f2fd;
                color: #1976d2;
            }}
        """)
        self.image_list.setAlternatingRowColors(True)
        image_list_layout.addWidget(self.image_list)
        
        right_layout.addWidget(image_list_group)
        
        main_layout.addWidget(right_panel, 1)
        
        self.tabs.addTab(image_widget, "画像管理")
        
    def setup_upload_tab(self):
        """アップロード管理タブ"""
        upload_widget = QWidget()
        layout = QVBoxLayout(upload_widget)
        
        # タブウィジェットでアップロード先を分ける
        upload_tabs = QTabWidget()
        
        # 楽天タブ
        rakuten_widget = QWidget()
        rakuten_layout = QVBoxLayout(rakuten_widget)
        self.setup_rakuten_upload(rakuten_layout)
        upload_tabs.addTab(rakuten_widget, "楽天市場")
        
        # Yahooタブ
        yahoo_widget = QWidget()
        yahoo_layout = QVBoxLayout(yahoo_widget)
        self.setup_yahoo_upload(yahoo_layout)
        upload_tabs.addTab(yahoo_widget, "Yahooショッピング")
        
        layout.addWidget(upload_tabs)
        self.tabs.addTab(upload_widget, "アップロード")
    
    def setup_rakuten_upload(self, layout):
        # 簡略化されたアップロードタブ
        main_layout = QVBoxLayout()
        
        # 画像フォルダ設定
        image_folder_group = QGroupBox("📁 画像フォルダ設定")
        image_folder_layout = QVBoxLayout(image_folder_group)
        
        self.rakuten_image_folder_label = QLabel("📂 未設定")
        self.rakuten_image_folder_label.setWordWrap(True)
        image_folder_layout.addWidget(self.rakuten_image_folder_label)
        
        browse_image_btn = QPushButton("📂 画像フォルダを選択")
        browse_image_btn.clicked.connect(self.browse_rakuten_image_folder)
        image_folder_layout.addWidget(browse_image_btn)
        
        main_layout.addWidget(image_folder_group)
        
        # 説明
        info_group = QGroupBox("アップロード手順")
        info_layout = QVBoxLayout(info_group)
        
        info_text = QLabel("1. 画像フォルダを選択\n"
                          "2. 外部FTPツールを使用してアップロード\n"
                          "   - CSV: /csv/ フォルダ\n"
                          "   - 画像: /cabinet/images/ フォルダ\n\n"
                          "※先に「商品情報入力」タブでCSVを生成してください")
        info_text.setWordWrap(True)
        info_layout.addWidget(info_text)
        
        main_layout.addWidget(info_group)
        main_layout.addStretch()
        
        layout.addLayout(main_layout)
        
    def setup_yahoo_upload(self, layout):
        """ヤフーアップロード設定"""
        # メインレイアウト（横分割）
        main_layout = QHBoxLayout()
        
        # 左側: コントロールパネル
        control_panel = QWidget()
        panel_width = int(350 * self.dpi_scale)
        control_panel.setMaximumWidth(panel_width)
        control_layout = QVBoxLayout(control_panel)
        
        # 店舗選択
        store_group = QGroupBox("店舗選択")
        store_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        store_layout = QVBoxLayout(store_group)
        
        self.yahoo_store_combo = QComboBox()
        self.yahoo_store_combo.addItems(["大宝家具 1号店 (taiho-kagu)", "大宝家具 2号店 (taiho-kagu2)"])
        combo_height = int(30 * self.dpi_scale)
        self.yahoo_store_combo.setMinimumHeight(combo_height)
        store_layout.addWidget(self.yahoo_store_combo)
        
        # URL表示
        self.yahoo_url_label = QLabel()
        self.yahoo_url_label.setStyleSheet("font-size: 10px; color: #666; padding: 5px;")
        self.yahoo_url_label.setWordWrap(True)
        self.update_yahoo_url()
        store_layout.addWidget(self.yahoo_url_label)
        
        control_layout.addWidget(store_group)
        
        # 操作ボタン
        action_group = QGroupBox("操作")
        action_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        action_layout = QVBoxLayout(action_group)
        
        load_page_btn = QPushButton("🌐 ストアクリエイターPro を読み込み")
        btn_height = int(35 * self.dpi_scale)
        btn_font_size = int(12 * self.dpi_scale)
        load_page_btn.setMinimumHeight(btn_height)
        load_page_btn.setStyleSheet(f"QPushButton {{ background-color: #FF9800; color: white; font-weight: bold; font-size: {btn_font_size}px; }}")
        load_page_btn.clicked.connect(self.load_yahoo_page)
        action_layout.addWidget(load_page_btn)
        
        prepare_yahoo_csv_btn = QPushButton("📁 Yahoo用CSVを準備")
        prepare_yahoo_csv_btn.setMinimumHeight(btn_height)
        prepare_yahoo_csv_btn.setStyleSheet(f"QPushButton {{ background-color: #4CAF50; color: white; font-weight: bold; font-size: {btn_font_size}px; }}")
        prepare_yahoo_csv_btn.clicked.connect(self.prepare_yahoo_csv)
        action_layout.addWidget(prepare_yahoo_csv_btn)
        
        open_folder_btn = QPushButton("📂 出力フォルダを開く")
        open_folder_btn.setMinimumHeight(btn_height)
        open_folder_btn.setStyleSheet(f"QPushButton {{ background-color: #2196F3; color: white; font-weight: bold; font-size: {btn_font_size}px; }}")
        open_folder_btn.clicked.connect(self.open_output_folder)
        action_layout.addWidget(open_folder_btn)
        
        control_layout.addWidget(action_group)
        
        # アップロード手順
        info_group = QGroupBox("アップロード手順")
        info_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        info_layout = QVBoxLayout(info_group)
        
        yahoo_info = QLabel("1. 店舗を選択\n"
                           "2. ストアクリエイターProを読み込み\n"
                           "3. Yahoo用CSVを準備\n"
                           "4. ログイン後、商品データアップロード\n"
                           "5. yahoo_item.csv をアップロード\n"
                           "6. オプションデータアップロード\n"
                           "7. yahoo_option.csv をアップロード")
        yahoo_info.setWordWrap(True)
        yahoo_info_font_size = max(12, int(14 * self.dpi_scale))
        yahoo_info.setStyleSheet(f"color: #666; font-size: {yahoo_info_font_size}px;")
        info_layout.addWidget(yahoo_info)
        
        control_layout.addWidget(info_group)
        control_layout.addStretch()
        
        # コンボボックスの変更を監視
        self.yahoo_store_combo.currentIndexChanged.connect(self.update_yahoo_url)
        self.yahoo_store_combo.currentIndexChanged.connect(self.load_yahoo_page)
        
        main_layout.addWidget(control_panel)
        
        # 右側: ブラウザ表示
        browser_group = QGroupBox("ストアクリエイターPro")
        browser_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        browser_layout = QVBoxLayout(browser_group)
        
        # ブラウザウィジェット
        self.yahoo_browser = QWebEngineView()
        browser_layout.addWidget(self.yahoo_browser)
        
        main_layout.addWidget(browser_group, 1)
        
        layout.addLayout(main_layout)
        
    def setup_check_tab(self):
        """ページ確認タブ"""
        check_widget = QWidget()
        layout = QHBoxLayout(check_widget)  # 横分割レイアウト
        
        # 左側: コントロールパネル
        left_panel = QWidget()
        panel_width = int(300 * self.dpi_scale)
        left_panel.setMaximumWidth(panel_width)
        left_layout = QVBoxLayout(left_panel)
        
        # 商品コード入力
        code_group = QGroupBox("🔢 商品コード入力")
        code_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        code_layout = QVBoxLayout(code_group)
        
        code_input_label = QLabel("商品コード:")
        code_input_label.setStyleSheet("font-weight: bold; margin-bottom: 5px;")
        code_layout.addWidget(code_input_label)
        
        self.product_code_input = QLineEdit()
        self.product_code_input.setPlaceholderText("10桁の商品コードを入力すると自動でURL生成")
        input_height = int(30 * self.dpi_scale)
        input_padding = int(8 * self.dpi_scale)
        input_font_size = int(12 * self.dpi_scale)
        input_border_radius = int(5 * self.dpi_scale)
        self.product_code_input.setMinimumHeight(input_height)
        self.product_code_input.setStyleSheet(f"""
            QLineEdit {{
                padding: {input_padding}px;
                border: 2px solid #ddd;
                border-radius: {input_border_radius}px;
                font-size: {input_font_size}px;
                background-color: white;
            }}
            QLineEdit:focus {{
                border-color: #2196F3;
                background-color: #f8f9fa;
            }}
        """)
        self.product_code_input.textChanged.connect(self.auto_generate_urls)  # 10桁で自動生成
        code_layout.addWidget(self.product_code_input)
        
        left_layout.addWidget(code_group)
        
        # URL選択リスト
        url_group = QGroupBox("🌐 ページ選択")
        url_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        url_layout = QVBoxLayout(url_group)
        
        self.url_list_widget = QListWidget()
        self.url_list_widget.itemClicked.connect(self.load_selected_page)
        list_font_size = int(11 * self.dpi_scale)
        list_padding = int(8 * self.dpi_scale)
        list_border_radius = int(3 * self.dpi_scale)
        self.url_list_widget.setStyleSheet(f"""
            QListWidget {{
                background-color: #fafafa;
                border: 1px solid #ddd;
                border-radius: {list_border_radius}px;
                font-size: {list_font_size}px;
            }}
            QListWidget::item {{
                padding: {list_padding}px;
                border-bottom: 1px solid #eee;
            }}
            QListWidget::item:selected {{
                background-color: #e3f2fd;
                color: #1976d2;
                font-weight: bold;
            }}
            QListWidget::item:hover {{
                background-color: #f5f5f5;
            }}
        """)
        url_layout.addWidget(self.url_list_widget)
        
        # URL手動入力
        manual_url_label = QLabel("直接URL入力:")
        manual_url_label.setStyleSheet("font-weight: bold; margin-top: 10px; margin-bottom: 5px;")
        url_layout.addWidget(manual_url_label)
        
        manual_url_layout = QHBoxLayout()
        self.manual_url_input = QLineEdit()
        self.manual_url_input.setPlaceholderText("https://...")
        manual_padding = int(6 * self.dpi_scale)
        manual_font_size = int(11 * self.dpi_scale)
        manual_border_radius = int(3 * self.dpi_scale)
        self.manual_url_input.setStyleSheet(f"""
            QLineEdit {{
                padding: {manual_padding}px;
                border: 1px solid #ddd;
                border-radius: {manual_border_radius}px;
                font-size: {manual_font_size}px;
            }}
            QLineEdit:focus {{
                border-color: #2196F3;
            }}
        """)
        manual_url_layout.addWidget(self.manual_url_input)
        
        load_manual_btn = QPushButton("🔗 読込")
        manual_btn_height = int(28 * self.dpi_scale)
        manual_btn_font_size = int(11 * self.dpi_scale)
        manual_btn_padding = int(10 * self.dpi_scale)
        load_manual_btn.setMinimumHeight(manual_btn_height)
        load_manual_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #FF9800;
                color: white;
                font-weight: bold;
                border-radius: {manual_border_radius}px;
                font-size: {manual_btn_font_size}px;
                padding: 0 {manual_btn_padding}px;
            }}
            QPushButton:hover {{
                background-color: #F57C00;
            }}
        """)
        load_manual_btn.clicked.connect(self.load_manual_url)
        manual_url_layout.addWidget(load_manual_btn)
        
        url_layout.addLayout(manual_url_layout)
        left_layout.addWidget(url_group)
        
        # ページ操作
        control_group = QGroupBox("🎮 ページ操作")
        control_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        control_layout = QVBoxLayout(control_group)
        
        nav_layout = QHBoxLayout()
        
        back_btn = QPushButton("⬅️ 戻る")
        nav_btn_height = int(30 * self.dpi_scale)
        nav_btn_font_size = int(11 * self.dpi_scale)
        nav_border_radius = int(3 * self.dpi_scale)
        back_btn.setMinimumHeight(nav_btn_height)
        back_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #607D8B;
                color: white;
                font-weight: bold;
                border-radius: {nav_border_radius}px;
                font-size: {nav_btn_font_size}px;
            }}
            QPushButton:hover {{
                background-color: #455A64;
            }}
        """)
        back_btn.clicked.connect(self.page_back)
        nav_layout.addWidget(back_btn)
        
        forward_btn = QPushButton("進む ➡️")
        forward_btn.setMinimumHeight(nav_btn_height)
        forward_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #607D8B;
                color: white;
                font-weight: bold;
                border-radius: {nav_border_radius}px;
                font-size: {nav_btn_font_size}px;
            }}
            QPushButton:hover {{
                background-color: #455A64;
            }}
        """)
        forward_btn.clicked.connect(self.page_forward)
        nav_layout.addWidget(forward_btn)
        
        reload_btn = QPushButton("🔄 更新")
        reload_btn.setMinimumHeight(nav_btn_height)
        reload_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                border-radius: {nav_border_radius}px;
                font-size: {nav_btn_font_size}px;
            }}
            QPushButton:hover {{
                background-color: #45a049;
            }}
        """)
        reload_btn.clicked.connect(self.page_reload)
        nav_layout.addWidget(reload_btn)
        
        control_layout.addLayout(nav_layout)
        left_layout.addWidget(control_group)
        
        # 確認結果
        result_group = QGroupBox("📋 確認結果")
        result_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        result_layout = QVBoxLayout(result_group)
        
        self.check_result = QTextEdit()
        self.check_result.setReadOnly(True)
        self.check_result.setMaximumHeight(120)
        self.check_result.setPlaceholderText("ページ読み込み結果がここに表示されます...")
        self.check_result.setStyleSheet("""
            QTextEdit {
                background-color: #fafafa;
                border: 1px solid #ddd;
                border-radius: 3px;
                font-family: 'Courier New', monospace;
                font-size: 10px;
                padding: 5px;
            }
        """)
        result_layout.addWidget(self.check_result)
        
        left_layout.addWidget(result_group)
        left_layout.addStretch()
        
        layout.addWidget(left_panel)
        
        # 右側: Webブラウザ表示
        browser_group = QGroupBox("🖥️ ページプレビュー")
        browser_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: {int(10 * self.dpi_scale)}px;
                padding-top: {int(10 * self.dpi_scale)}px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {int(10 * self.dpi_scale)}px;
                padding: 0 {int(5 * self.dpi_scale)}px 0 {int(5 * self.dpi_scale)}px;
            }}
        """)
        browser_layout = QVBoxLayout(browser_group)
        
        # URL表示バー
        url_bar_layout = QHBoxLayout()
        url_icon = QLabel("🌐")
        url_icon.setStyleSheet("font-size: 12px; margin-right: 5px;")
        url_bar_layout.addWidget(url_icon)
        
        self.current_url_label = QLabel("未選択")
        self.current_url_label.setStyleSheet("""
            QLabel {
                font-size: 10px;
                color: #666;
                padding: 5px 8px;
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 3px;
                font-family: 'Courier New', monospace;
            }
        """)
        self.current_url_label.setWordWrap(False)
        self.current_url_label.setMaximumHeight(25)
        url_bar_layout.addWidget(self.current_url_label, 1)
        
        browser_layout.addLayout(url_bar_layout)
        
        # ブラウザウィジェット
        self.page_browser = QWebEngineView()
        browser_border_radius = int(3 * self.dpi_scale)
        self.page_browser.setStyleSheet(f"""
            QWebEngineView {{
                border: 1px solid #ddd;
                border-radius: {browser_border_radius}px;
                background-color: white;
            }}
        """)
        
        # JavaScriptコンソールエラーを無視する設定
        from PyQt5.QtWebEngineWidgets import QWebEngineSettings
        settings = self.page_browser.settings()
        settings.setAttribute(QWebEngineSettings.JavascriptEnabled, True)
        settings.setAttribute(QWebEngineSettings.ErrorPageEnabled, False)
        
        browser_layout.addWidget(self.page_browser)
        
        layout.addWidget(browser_group, 1)  # 最大領域を占有
        
        self.tabs.addTab(check_widget, "ページ確認")
        
    def setup_logging(self):
        """ログ設定"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
    def log_message(self, message, level="INFO"):
        """ログメッセージ表示"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {level}: {message}"
        if hasattr(self, 'log_text'):
            self.log_text.append(log_entry)
        # ステータスバーへの表示は必要な時のみ行う
        # self.statusBar().showMessage(message)
        
    def save_ftp_password(self):
        """FTPパスワードを保存（簡易暗号化）"""
        password = self.ftp_pass_input.text()
        if password:
            try:
                # 簡単な暗号化
                encoded = base64.b64encode(password.encode()).decode()
                config = configparser.ConfigParser()
                config['FTP'] = {'password': encoded}
                
                config_path = Path(__file__).parent / '.config.ini'
                with open(config_path, 'w') as f:
                    config.write(f)
                    
                QMessageBox.information(self, "保存完了", "パスワードが保存されました。")
            except Exception as e:
                QMessageBox.warning(self, "保存エラー", f"パスワードの保存に失敗しました: {str(e)}")
        else:
            QMessageBox.warning(self, "入力エラー", "パスワードを入力してください。")
            
    def load_ftp_password(self):
        """保存されたFTPパスワードを読み込み"""
        try:
            config_path = Path(__file__).parent / '.config.ini'
            if config_path.exists():
                config = configparser.ConfigParser()
                config.read(config_path)
                
                if 'FTP' in config and 'password' in config['FTP']:
                    encoded = config['FTP']['password']
                    password = base64.b64decode(encoded).decode()
                    if hasattr(self, 'ftp_pass_input'):
                        self.ftp_pass_input.setText(password)
        except Exception:
            pass  # エラーがある場合は何もしない
        
    def browse_master_tool(self):
        """マスタツール選択"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "HANBAIMENU.exe を選択", "", "実行ファイル (*.exe)"
        )
        if file_path:
            self.master_tool_path = file_path
            self.master_path_label.setText(os.path.basename(file_path))
            self.save_settings()
            
    def auto_launch_master(self):
        """マスタツール自動起動・埋め込み"""
        self.launch_and_embed_master()
    
    def launch_and_embed_master(self):
        """マスタツール起動・埋め込み"""
        if not self.master_tool_path:
            QMessageBox.warning(self, "警告", "マスタツールのパスが設定されていません")
            return
            
        # ネットワークパスの確認
        if not os.path.exists(self.master_tool_path):
            error_msg = f"マスタツール '{self.master_tool_path}' が見つかりません。\n\n考えられる原因:\n1. ネットワークドライブが接続されていない\n2. ファイルパスが間違っている\n3. アクセス権限がない\n\nネットワーク接続を確認してください。"
            QMessageBox.warning(self, "マスタツール接続エラー", error_msg)
            print(f"マスタツールパスエラー: {self.master_tool_path}")
            return
            
        try:
            # 起動前に既存のプロセスをクリーンアップ
            import subprocess
            import time
            try:
                result = subprocess.run(['taskkill', '/f', '/im', 'HAMST040.exe'], 
                                     capture_output=True, text=True)
                if result.returncode == 0:
                    print("既存のHAMST040プロセスをクリーンアップしました")
                    time.sleep(1)  # 1秒待機
            except:
                pass
            
            # 作業ディレクトリをDBConfig.xmlがある場所に設定
            work_dir = os.path.dirname(self.master_tool_path)
            print(f"マスタツール起動: {self.master_tool_path}")
            print(f"作業ディレクトリ: {work_dir}")
            
            # HAMST040.exeを起動
            self.master_process = subprocess.Popen(
                self.master_tool_path, 
                cwd=work_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            print(f"プロセスID: {self.master_process.pid}")
            
            # プロセス起動確認のため少し待つ
            time.sleep(2)
            
            # プロセスが生きているか確認
            if self.master_process.poll() is not None:
                # プロセスが終了している場合、エラー情報を取得
                stdout, stderr = self.master_process.communicate()
                error_msg = f"マスタツールが起動直後に終了しました。\n"
                error_msg += f"終了コード: {self.master_process.returncode}\n"
                if stderr:
                    error_msg += f"エラー出力: {stderr.decode('utf-8', errors='ignore')}\n"
                if stdout:
                    error_msg += f"標準出力: {stdout.decode('utf-8', errors='ignore')}\n"
                print(error_msg)
                QMessageBox.warning(self, "マスタツール起動エラー", error_msg)
                return
            else:
                print("プロセスは正常に動作中")
            
            # 埋め込み試行回数をリセット
            self.embed_attempt_count = 0
            
            # 埋め込みを試行するタイマーを開始
            self.embed_timer.start(1000)  # 1秒ごとにチェック（ウィンドウ表示に時間がかかる場合があるため）
            
        except Exception as e:
            error_msg = f"起動に失敗しました: {str(e)}"
            print(f"エラー詳細: {error_msg}")
            import traceback
            print(f"スタックトレース: {traceback.format_exc()}")
            QMessageBox.critical(self, "エラー", error_msg)
    
    def try_embed_master(self):
        """マスタツールの埋め込みを試行"""
        try:
            self.embed_attempt_count += 1
            
            # プロセスが生きているかチェック
            if self.master_process and self.master_process.poll() is not None:
                self.embed_timer.stop()
                stdout, stderr = self.master_process.communicate()
                error_msg = f"マスタツールプロセスが終了しました。\n"
                error_msg += f"終了コード: {self.master_process.returncode}\n"
                if stderr:
                    error_msg += f"エラー: {stderr.decode('utf-8', errors='ignore')[:200]}\n"
                print(error_msg)
                QMessageBox.warning(self, "マスタツールエラー", error_msg)
                return
            
            # 30回試行後（30秒後）にタイムアウト
            if self.embed_attempt_count > 30:
                self.embed_timer.stop()
                reply = QMessageBox.question(self, "マスタツール検出", 
                    "マスタツール「商品一覧」ウィンドウが自動検出できませんでした。\n"
                    "手動でウィンドウを選択しますか？",
                    QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.Yes:
                    self.manual_embed_window()
                return
            
            # 全ウィンドウを検索してデバッグ情報を表示
            found_windows = []
            
            def enum_windows(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    window_title = win32gui.GetWindowText(hwnd)
                    if window_title:  # 空でないタイトルのみ
                        found_windows.append(window_title)
                        # HAMST040.exeのウィンドウを検索（実際のウィンドウタイトルに基づく）
                        if window_title == "商品一覧" or any(keyword in window_title for keyword in ["HAMST", "マスタ", "大宝", "TAIHO"]):
                            # self.log_message(f"候補ウィンドウを発見: {window_title}")
                            self.master_hwnd = hwnd
                            return False
                return True
            
            win32gui.EnumWindows(enum_windows, None)
            
            # デバッグ情報をログに出力
            if not self.master_hwnd:
                print(f"検索キーワード: HAMST, マスタ, 大宝, TAIHO, 040")
                # print(f"現在開いているウィンドウ: {found_windows}") # ログが長くなるためコメントアウト
                # self.log_message(f"自動検出失敗。現在開いているウィンドウ: {', '.join(found_windows[:10])}")  # 最初の10個だけ表示
            
            if self.master_hwnd:
                try:
                    window_title = win32gui.GetWindowText(self.master_hwnd)
                    print(f"ウィンドウを発見: {window_title} (HWND: {self.master_hwnd})")

                    parent_hwnd = int(self.master_embed_widget.winId())
                    print(f"埋め込み先ウィジェットの HWND: {parent_hwnd}")
                    
                    # 親ウィンドウを設定
                    win32gui.SetParent(self.master_hwnd, parent_hwnd)
                    
                    # ウィンドウスタイルを調整 (タイトルバーなどを削除し、子ウィンドウスタイルを設定)
                    style = win32gui.GetWindowLong(self.master_hwnd, win32con.GWL_STYLE)
                    # WS_OVERLAPPEDWINDOW スタイル (WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX) を除去
                    # WS_POPUP スタイルも除去する可能性を考慮
                    style &= ~(win32con.WS_CAPTION | win32con.WS_THICKFRAME | win32con.WS_SYSMENU | win32con.WS_MINIMIZEBOX | win32con.WS_MAXIMIZEBOX)
                    style |= win32con.WS_CHILD # 子ウィンドウスタイルを追加
                    win32gui.SetWindowLong(self.master_hwnd, win32con.GWL_STYLE, style)

                    # ウィンドウの状態を通常に戻す試み
                    win32gui.ShowWindow(self.master_hwnd, win32con.SW_RESTORE)
                    QApplication.processEvents() # OSに状態変更を処理させる
                    # time.sleep(0.05) # 必要に応じて短い待機


                    # 埋め込みウィジェットのサイズを取得
                    widget_size = self.master_embed_widget.size()
                    target_width = widget_size.width()
                    target_height = widget_size.height()
                    print(f"埋め込みウィジェットのサイズ: {target_width}x{target_height}")

                    # ウィンドウを表示域内に移動・リサイズ (0,0 は親ウィジェットの左上からの相対位置)
                    # SWP_FRAMECHANGED を追加してスタイル変更を適用
                    win32gui.SetWindowPos(
                        self.master_hwnd,
                        0,  # HWND_TOP or specific z-order (0 or win32con.HWND_TOP)
                        0, 0, # x, y (親ウィジェットのクライアント座標系の左上)
                        target_width, target_height,
                        win32con.SWP_SHOWWINDOW | win32con.SWP_FRAMECHANGED | win32con.SWP_NOACTIVATE
                    )
                    
                    self.embed_timer.stop()
                    print("表示域内配置完了")
                    self.last_resize_size = (target_width, target_height) # 初期サイズを記録
                    # 初期表示を確実にするために、少し遅れてリサイズを再試行 (遅延を250msに延長)
                    QTimer.singleShot(250, self.resize_embedded_window)
                    
                except Exception as embed_error:
                    self.log_message(f"埋め込み処理エラー: {str(embed_error)}", "ERROR")
                
        except Exception as e:
            self.log_message(f"ウィンドウ検索エラー: {str(e)}", "WARNING")
    
    def close_master_tool(self):
        """マスタツール終了"""
        try:
            # すべてのHAMST040プロセスを確実に終了
            import subprocess
            try:
                subprocess.run(['taskkill', '/f', '/im', 'HAMST040.exe'], 
                             capture_output=True, text=True)
                print("すべてのHAMST040プロセスを強制終了しました")
            except:
                pass
            
            if self.master_process:
                try:
                    self.master_process.terminate()
                    self.master_process.wait(timeout=2)
                except:
                    try:
                        self.master_process.kill()
                    except:
                        pass
                self.master_process = None
            
            if self.master_hwnd:
                try:
                    win32gui.CloseWindow(self.master_hwnd)
                except:
                    pass
                self.master_hwnd = None
            
            # 状態をリセット
            self.embed_timer.stop()
            self.resize_timer.stop()
            self.embed_attempt_count = 0
            self.last_resize_size = None
            
            self.log_message("マスタツールを完全終了しました")
            
        except Exception as e:
            print(f"終了エラー: {str(e)}")
            QMessageBox.critical(self, "エラー", f"終了に失敗しました: {str(e)}")
    
    def manual_embed_window(self):
        """手動でウィンドウを選択して埋め込み"""
        try:
            # 現在開いているウィンドウを取得
            windows = []
            
            def enum_windows(hwnd, _):
                if win32gui.IsWindowVisible(hwnd):
                    window_title = win32gui.GetWindowText(hwnd)
                    if window_title and len(window_title) > 1:  # 意味のあるタイトルのみ
                        windows.append((hwnd, window_title))
                return True
            
            win32gui.EnumWindows(enum_windows, None)
            
            if not windows:
                QMessageBox.information(self, "情報", "埋め込み可能なウィンドウが見つかりません")
                return
            
            # ウィンドウ選択ダイアログ
            window_titles = [title for _, title in windows]
            selected_title, ok = QInputDialog.getItem(
                self, "ウィンドウ選択", 
                "埋め込みたいウィンドウを選択してください:", 
                window_titles, 0, False
            )
            
            if ok and selected_title:
                # 選択されたウィンドウのハンドルを取得
                selected_hwnd = None
                for hwnd, title in windows:
                    if title == selected_title:
                        selected_hwnd = hwnd
                        break
                
                if selected_hwnd:
                    self.master_hwnd = selected_hwnd
                    self.embed_selected_window()
                    
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"手動埋め込みに失敗しました: {str(e)}")
    
    def embed_selected_window(self):
        """選択されたウィンドウを埋め込み"""
        try:
            if not self.master_hwnd:
                return
            
            parent_hwnd = int(self.master_embed_widget.winId())
            window_title = win32gui.GetWindowText(self.master_hwnd)
            
            self.log_message(f"手動埋め込み開始: {window_title}")
            
            # 親ウィンドウを設定
            result = win32gui.SetParent(self.master_hwnd, parent_hwnd)
            if result == 0:
                self.log_message("SetParent failed", "WARNING")
                return
            
            # ウィンドウスタイルを調整
            style = win32gui.GetWindowLong(self.master_hwnd, win32con.GWL_STYLE)
            style = style & ~win32con.WS_CAPTION & ~win32con.WS_THICKFRAME
            win32gui.SetWindowLong(self.master_hwnd, win32con.GWL_STYLE, style)
            
            # ウィンドウサイズを表示域目一杯に調整
            widget_size = self.master_embed_widget.size()
            target_width = widget_size.width()
            target_height = widget_size.height()
            
            win32gui.SetWindowPos(
                self.master_hwnd, 0, 
                0, 0,  # 親ウィジェットのクライアント座標 (0,0)
                target_width, target_height,
                win32con.SWP_SHOWWINDOW | win32con.SWP_FRAMECHANGED | win32con.SWP_NOACTIVATE
            )
            
            # 初期サイズを記録
            self.last_resize_size = (target_width, target_height)
            # 初期表示を確実にするために、少し遅れてリサイズを再試行 (遅延を250msに延長)
            QTimer.singleShot(250, self.resize_embedded_window)
            
            # self.log_message(f"手動埋め込み完了: {window_title}")
            
        except Exception as e:
            self.log_message(f"手動埋め込みエラー: {str(e)}", "ERROR")
    
    def on_embed_widget_resize(self, event):
        """埋め込みウィジェットリサイズ時の処理"""
        # 現在のサイズを取得
        current_size = (self.master_embed_widget.width(), self.master_embed_widget.height())
        print(f"埋め込みウィジェットリサイズ検出: {current_size}")
        
        # サイズが小さすぎる場合はスキップ（初期化中など）
        if current_size[0] < 100 or current_size[1] < 100:
            if hasattr(event, 'accept'):
                event.accept()
            return
        
        # 常にリサイズを実行（前回と同じサイズでも）
        # タイマーをリセットして、リサイズが完了してから実行
        self.resize_timer.stop()
        self.resize_timer.start(300)  # 300ms後に実行
        if hasattr(event, 'accept'):
            event.accept()
    
    def resize_embedded_window(self):
        """埋め込まれたウィンドウのサイズを調整"""
        if not self.master_hwnd:
            return
            
        try:
            parent_rect = self.master_embed_widget.geometry()
            current_size = (parent_rect.width(), parent_rect.height())
            
            # サイズが小さすぎる場合はスキップ
            if current_size[0] < 100 or current_size[1] < 100:
                return
            
            # ウィンドウが最小化されていないかチェック
            if not win32gui.IsWindowVisible(self.master_hwnd):
                return
            
            win32gui.SetWindowPos(
                self.master_hwnd, 0, 0, 0, 
                current_size[0],
                current_size[1],
                win32con.SWP_SHOWWINDOW | win32con.SWP_FRAMECHANGED | win32con.SWP_NOACTIVATE
            )
            
            # 前回のサイズを更新
            self.last_resize_size = current_size
            
        except Exception as e:
            self.log_message(f"リサイズエラー: {str(e)}", "WARNING")
            
    def resizeEvent(self, event):
        """メインウィンドウリサイズ時の処理"""
        super().resizeEvent(event)
        # マスタタブが表示されている場合のみリサイズ
        if self.tabs.currentIndex() == 1 and self.master_hwnd:
            print("メインウィンドウリサイズ検出")
            # リサイズが完了してから実行（頻繁な実行を防ぐ）
            self.resize_timer.stop()
            self.resize_timer.start(200)  # より短い間隔で実行
    
    def manual_resize_master(self):
        """手動リサイズボタンが押された時の処理"""
        if self.master_hwnd:
            # もう埋め込みは諦めて、外部ウィンドウとして最大化
            try:
                print("埋め込みを諦めて、ウィンドウを最大化します")
                
                # ウィンドウを最大化
                win32gui.ShowWindow(self.master_hwnd, win32con.SW_MAXIMIZE)
                
                # 統合ツール自体を最小化して邪魔にならないようにする
                self.showMinimized()
                
                self.log_message("マスタツールを最大化しました。統合ツールは最小化されました。", "INFO")
                
            except Exception as e:
                print(f"最大化エラー: {e}")
                self.log_message(f"最大化エラー: {str(e)}", "ERROR")
                
        else:
            print("埋め込みウィンドウハンドルがNullです")
            self.log_message("埋め込みウィンドウが見つかりません", "WARNING")
    
    def on_tab_changed(self, index):
        """タブが切り替わった時の処理"""
        # マスタタブ（index 1）に切り替わった場合
        if index == 1 and self.master_hwnd:
            # サイズ調整のみ行う
            QTimer.singleShot(100, self.resize_embedded_window)  # 少し遅延させて確実に実行
            
    def browse_rakuten_image_folder(self):
        """楽天用画像フォルダ選択"""
        folder = QFileDialog.getExistingDirectory(self, "画像フォルダを選択")
        if folder:
            self.rakuten_image_folder_label.setText(f"📂 {folder}")
            self.rakuten_image_folder = folder
            
    def browse_image_folder(self):
        """画像フォルダ選択（互換性のため残す）"""
        folder = QFileDialog.getExistingDirectory(self, "画像フォルダを選択")
        if folder:
            self.image_folder_label.setText(f"📂 {folder}")
            self.refresh_image_list()
            
    def refresh_image_list(self):
        """画像リスト更新"""
        folder_text = self.image_folder_label.text()
        # "📂 " プレフィックスを除去
        folder = folder_text.replace("📂 ", "") if folder_text.startswith("📂 ") else folder_text
        
        if folder and folder != "未設定":
            self.image_list.clear()
            image_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp')
            count = 0
            
            try:
                folder_path = Path(folder)
                if folder_path.exists():
                    for file in folder_path.iterdir():
                        if file.suffix.lower() in image_extensions:
                            # ファイルサイズも表示
                            file_size = file.stat().st_size
                            size_mb = file_size / (1024 * 1024)
                            self.image_list.addItem(f"🖼️ {file.name} ({size_mb:.1f}MB)")
                            count += 1
                    
                    self.image_count_label.setText(f"画像ファイル数: {count}")
                    self.log_message(f"画像リストを更新: {count}ファイル")
                else:
                    self.image_count_label.setText("画像ファイル数: フォルダが見つかりません")
            except Exception as e:
                self.log_message(f"画像リスト更新エラー: {str(e)}", "WARNING")
                self.image_count_label.setText("画像ファイル数: エラー")
    
    def open_current_image_folder(self):
        """現在設定されている画像フォルダを開く"""
        folder_text = self.image_folder_label.text()
        folder = folder_text.replace("📂 ", "") if folder_text.startswith("📂 ") else folder_text
        
        if folder and folder != "未設定":
            try:
                if sys.platform == "win32":
                    os.startfile(folder)
                else:
                    subprocess.call(["open" if sys.platform == "darwin" else "xdg-open", folder])
                self.log_message(f"フォルダを開きました: {folder}")
            except Exception as e:
                self.log_message(f"フォルダを開けませんでした: {str(e)}", "WARNING")
                    
    def connect_sftp_with_retry(self):
        """SFTPに接続（パスワード自動切り替え）"""
        passwords = ["ta1hoKa9", "Ta1hoka9"]
        
        if self.ftp_pass_input.text():
            # ユーザーが入力したパスワードを優先
            passwords.insert(0, self.ftp_pass_input.text())
        
        for password in passwords:
            try:
                transport = paramiko.Transport((self.ftp_server_input.text(), 22))
                transport.connect(username=self.ftp_user_input.text(), password=password)
                sftp = paramiko.SFTPClient.from_transport(transport)
                self.log_message(f"SFTP接続成功")
                return sftp, transport
            except Exception as e:
                self.log_message(f"パスワード {password} で接続失敗: {str(e)}", "WARNING")
                continue
        
        raise Exception("すべてのパスワードで接続に失敗しました")
    
    def upload_csv_to_rakuten(self):
        """楽天へCSVアップロード"""
        try:
            # CSV出力フォルダを確認
            csv_folder = os.path.join(os.path.dirname(__file__), "CSVTOOL")
            if not os.path.exists(csv_folder):
                QMessageBox.warning(self, "警告", "CSVフォルダが見つかりません。先にCSVを生成してください。")
                return
            
            # 最新のCSVフォルダを取得
            csv_dirs = [d for d in os.listdir(csv_folder) if os.path.isdir(os.path.join(csv_folder, d))]
            if not csv_dirs:
                QMessageBox.warning(self, "警告", "CSVファイルが見つかりません")
                return
            
            latest_dir = os.path.join(csv_folder, sorted(csv_dirs)[-1])
            
            # SFTP接続
            sftp, transport = self.connect_sftp_with_retry()
            
            try:
                # CSVファイルをアップロード（順番通り）
                csv_files = [
                    ("rakuten_normal-item.csv", "normal-item.csv"),
                    ("rakuten_item-cat.csv", "item-cat.csv")
                ]
                
                sftp.chdir("/ritem/batch")
                
                for local_name, remote_name in csv_files:
                    local_path = os.path.join(latest_dir, local_name)
                    if os.path.exists(local_path):
                        self.log_message(f"{local_name} をアップロード中...")
                        # ファイルサイズを確認
                        file_size = os.path.getsize(local_path)
                        uploaded = 0
                        
                        # プログレスバー更新用のコールバック
                        def progress_callback(transferred, total):
                            percent = int((transferred / total) * 100)
                            self.progress_bar.setValue(percent)
                            QApplication.processEvents()
                        
                        # ファイルをアップロード
                        with open(local_path, 'rb') as f:
                            sftp.putfo(f, remote_name, callback=progress_callback)
                        
                        self.log_message(f"{remote_name} としてアップロード完了")
                    else:
                        self.log_message(f"{local_name} が見つかりません", "WARNING")
                
                self.log_message("CSVアップロードが完了しました")
                
            except Exception as e:
                # エラー時に再接続を試みる
                if "Connection" in str(e) or "timed out" in str(e):
                    self.log_message("接続が切れました。再接続します...", "WARNING")
                    sftp.close()
                    transport.close()
                    # 再度接続して続きから実行
                    self.upload_csv_to_rakuten()
                    return
                else:
                    raise e
            finally:
                sftp.close()
                transport.close()
                
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"CSVアップロードに失敗しました: {str(e)}")
    
    def upload_images_to_rakuten(self):
        """楽天へ画像アップロード"""
        try:
            # 楽天用画像フォルダをチェック
            if not hasattr(self, 'rakuten_image_folder') or not self.rakuten_image_folder:
                QMessageBox.warning(self, "警告", "画像フォルダが設定されていません")
                return
            
            # SFTP接続
            sftp, transport = self.connect_sftp_with_retry()
            
            try:
                sftp.chdir("/cabinet/images")
                
                # 画像ファイルをアップロード
                image_folder = self.rakuten_image_folder
                image_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp')
                uploaded_count = 0
                
                for file in Path(image_folder).iterdir():
                    if file.suffix.lower() in image_extensions:
                        self.log_message(f"{file.name} をアップロード中...")
                        sftp.put(str(file), file.name)
                        uploaded_count += 1
                
                self.log_message(f"画像アップロード完了: {uploaded_count}ファイル")
                
            finally:
                sftp.close()
                transport.close()
                
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"画像アップロードに失敗しました: {str(e)}")
            
    def update_yahoo_url(self):
        """選択された店舗のURLを更新"""
        if self.yahoo_store_combo.currentIndex() == 0:
            # 1号店
            url = "https://editor.store.yahoo.co.jp/RT/taiho-kagu/ItemMgr/index"
        else:
            # 2号店  
            url = "https://editor.store.yahoo.co.jp/RT/taiho-kagu2/ItemMgr/index"
        
        self.yahoo_url_label.setText(f'現在の店舗: {url}')
        return url
    
    def load_yahoo_page(self):
        """選択された店舗のページを読み込み"""
        url = self.update_yahoo_url()
        self.yahoo_browser.load(QUrl(url))
        self.log_message(f"Yahooショッピング {self.yahoo_store_combo.currentText()} を読み込みました")
    
    def prepare_yahoo_csv(self):
        """Yahoo用CSVファイルを確認してフォルダを開く"""
        try:
            # CSVフォルダを確認
            csv_folder = os.path.join(os.path.dirname(__file__), "CSVTOOL")
            if not os.path.exists(csv_folder):
                csv_folder = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'CSVTOOL')
            
            if not os.path.exists(csv_folder):
                QMessageBox.warning(self, "警告", "CSVフォルダが見つかりません")
                return
            
            # 最新のCSVフォルダを取得
            csv_dirs = [d for d in os.listdir(csv_folder) if os.path.isdir(os.path.join(csv_folder, d))]
            if not csv_dirs:
                QMessageBox.warning(self, "警告", "CSVファイルが見つかりません")
                return
            
            latest_dir = os.path.join(csv_folder, sorted(csv_dirs)[-1])
            
            # 出力されたCSVファイルを確認
            csv_files = []
            if self.yahoo_store_combo.currentIndex() == 0:
                # 1号店とヤフオク
                csv_files = ["yahoo_item.csv", "yahoo_option.csv", "yahoo_auction_item.csv", "yahoo_auction_option.csv"]
            else:
                # 2号店
                csv_files = ["yahoo2_item.csv", "yahoo2_option.csv"]
            
            # ファイル存在確認
            missing_files = []
            for csv_file in csv_files:
                if not os.path.exists(os.path.join(latest_dir, csv_file)):
                    missing_files.append(csv_file)
            
            if missing_files:
                self.log_message(f"警告: {', '.join(missing_files)} が見つかりません", "WARNING")
            else:
                self.log_message(f"CSVファイル確認完了: {', '.join(csv_files)}")
            
            # フォルダを開く
            if sys.platform == "win32":
                os.startfile(latest_dir)
                
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"CSV確認に失敗しました: {str(e)}")
    
    def open_output_folder(self):
        """出力フォルダを開く"""
        # デスクトップのCSVTOOLフォルダを開く
        output_path = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'CSVTOOL')
        if not os.path.exists(output_path):
            output_path = os.path.join(os.path.dirname(__file__), "output")
        
        if sys.platform == "win32":
            os.startfile(output_path)
        else:
            subprocess.call(["open", output_path])
            
    def check_pages(self):
        """ページ確認"""
        urls = self.url_list.toPlainText().strip().split('\n')
        self.check_result.clear()
        
        for url in urls:
            if url.strip():
                # TODO: 実際のページチェック処理
                self.check_result.append(f"✓ {url} - OK")
                
    def auto_execute_workflow(self):
        """ワークフロー自動実行"""
        self.log_message("ワークフローの自動実行を開始します")
        # TODO: 各ステップの自動実行
        
    def open_settings(self):
        """設定画面"""
        QMessageBox.information(self, "設定", "設定機能は開発中です")
        
    def show_help(self):
        """ヘルプ表示"""
        help_text = """
        統合EC管理ツール ヘルプ
        
        1. マスタ作成: HANBAIMENU.exeを使用してマスタを作成
        2. 商品情報入力: 商品データを入力・編集
        3. 画像管理: 商品画像を整理・管理
        4. アップロード: 各モールへデータをアップロード
        5. ページ確認: 作成したページの確認
        
        各機能の詳細は、それぞれのタブで確認してください。
        """
        QMessageBox.information(self, "ヘルプ", help_text)
        
    def save_settings(self):
        """設定保存"""
        settings = {
            "master_tool_path": self.master_tool_path,
            "ftp_server": self.ftp_server_input.text(),
            "ftp_user": self.ftp_user_input.text()
        }
        with open("integrated_tool_settings.json", "w") as f:
            json.dump(settings, f)
            
    def auto_generate_urls(self):
        """商品コード入力時の自動URL生成"""
        product_code = self.product_code_input.text().strip()
        
        # 10桁でない場合はリストをクリア
        if len(product_code) != 10:
            self.url_list_widget.clear()
            return
        
        # 数字以外が含まれている場合もリセット
        if not product_code.isdigit():
            self.url_list_widget.clear()
            return
            
        # 10桁の数字なら自動でURL生成
        urls = [
            f"https://item.rakuten.co.jp/taiho-kagu/{product_code}/",
            f"https://store.shopping.yahoo.co.jp/taiho-kagu/{product_code}.html", 
            f"https://store.shopping.yahoo.co.jp/taiho-kagu2/{product_code}.html"
        ]
        
        # URLリストウィジェットに追加
        self.url_list_widget.clear()
        for i, url in enumerate(urls):
            store_name = ["楽天市場", "Yahoo 1号店", "Yahoo 2号店"][i]
            self.url_list_widget.addItem(f"{store_name}: {url}")
        
        self.log_message(f"商品コード {product_code} のURLを生成しました")
        
        # 最初のURLを自動読み込み
        if urls:
            self.load_page_by_url(urls[0])
    
    def load_selected_page(self, item):
        """選択されたページを読み込み"""
        text = item.text()
        # "店舗名: URL" から URL部分を抽出
        if ": " in text:
            url = text.split(": ", 1)[1]
            self.load_page_by_url(url)
    
    def load_manual_url(self):
        """手動入力されたURLを読み込み"""
        url = self.manual_url_input.text().strip()
        if url:
            self.load_page_by_url(url)
        
    def load_page_by_url(self, url):
        """指定されたURLをブラウザに読み込み"""
        try:
            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url
            
            self.page_browser.load(QUrl(url))
            # URLを短縮表示
            display_url = url if len(url) <= 60 else url[:57] + "..."
            self.current_url_label.setText(f"URL: {display_url}")
            self.log_message(f"ページを読み込み中: {url}")
            
            # ページ読み込み完了を監視
            self.page_browser.loadFinished.connect(self.on_page_loaded)
            
        except Exception as e:
            self.log_message(f"ページ読み込みエラー: {str(e)}", "ERROR")
    
    def on_page_loaded(self, success):
        """ページ読み込み完了時の処理"""
        if success:
            current_url = self.page_browser.url().toString()
            self.check_result.append(f"✓ {current_url} - 読み込み成功")
            self.log_message("ページ読み込み完了")
        else:
            current_url = self.page_browser.url().toString()
            self.check_result.append(f"✗ {current_url} - 読み込み失敗")
            self.log_message("ページ読み込み失敗", "WARNING")
    
    def page_back(self):
        """ブラウザで戻る"""
        self.page_browser.back()
    
    def page_forward(self):
        """ブラウザで進む"""
        self.page_browser.forward()
    
    def page_reload(self):
        """ブラウザで更新"""
        self.page_browser.reload()
    
    
    def auto_execute_workflow(self):
        """ワークフロー自動実行"""
        self.log_message("ワークフローの自動実行を開始します")
        self.progress_bar.setValue(0)
        
        try:
            # 1. マスタ作成ツール起動
            if self.master_tool_path and os.path.exists(self.master_tool_path):
                self.log_message("1. マスタ作成ツールを起動...")
                work_dir = os.path.dirname(self.master_tool_path)
                subprocess.Popen(self.master_tool_path, cwd=work_dir)
                QMessageBox.information(self, "確認", "マスタ作成が完了したらOKを押してください")
                self.progress_bar.setValue(15)
            
            # 2. CSV生成
            self.log_message("2. CSV生成を実行...")
            self.generate_csv()
            self.progress_bar.setValue(30)
            
            # 3. 画像準備確認
            self.log_message("3. 画像準備を確認...")
            if not self.image_folder_label.text() or self.image_folder_label.text() == "未設定":
                QMessageBox.warning(self, "警告", "画像フォルダを設定してください")
                return
            self.progress_bar.setValue(45)
            
            # 4. 楽天アップロード
            reply = QMessageBox.question(self, "確認", "楽天市場へアップロードしますか？")
            if reply == QMessageBox.Yes:
                self.log_message("4. 楽天市場へアップロード...")
                self.upload_csv_to_rakuten()
                self.upload_images_to_rakuten()
            self.progress_bar.setValue(70)
            
            # 5. Yahoo準備
            self.log_message("5. Yahoo用ファイル準備...")
            self.prepare_yahoo_csv()
            self.progress_bar.setValue(85)
            
            # 6. 完了
            self.progress_bar.setValue(100)
            self.log_message("ワークフローが完了しました")
            QMessageBox.information(self, "完了", "全ての処理が完了しました")
            
        except Exception as e:
            self.log_message(f"エラー: {str(e)}", "ERROR")
            QMessageBox.critical(self, "エラー", f"処理中にエラーが発生しました: {str(e)}")
    
    def load_settings(self):
        """設定読み込み"""
        try:
            with open("integrated_tool_settings.json", "r") as f:
                settings = json.load(f)
                saved_path = settings.get("master_tool_path")
                if saved_path and os.path.exists(saved_path):
                    self.master_tool_path = saved_path
                    self.master_path_label.setText(self.master_tool_path)
                self.ftp_server_input.setText(settings.get("ftp_server", "upload.rakuten.ne.jp"))
                self.ftp_user_input.setText(settings.get("ftp_user", "taiho-kagu"))
        except:
            pass

def main():
    # 高DPI対応（QApplication作成前に設定）
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    
    # DPIスケールファクターを取得してフォントサイズを調整
    screen = app.primaryScreen()
    dpi = screen.physicalDotsPerInch()
    scale_factor = dpi / 96.0  # 96 DPIを基準とする
    
    # フォントサイズをスケールファクターに応じて調整
    font = app.font()
    base_font_size = 9  # ベースフォントサイズ
    scaled_font_size = max(8, int(base_font_size * scale_factor))
    font.setPointSize(scaled_font_size)
    app.setFont(font)
    
    window = IntegratedECTool()
    window.showMaximized()  # デフォルトで最大化表示
    window.load_settings()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()