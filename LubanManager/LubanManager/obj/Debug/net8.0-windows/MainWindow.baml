   M S B A M L   `   `   `  ����  LubanManager!    LubanManagerTool.MainWindow   5   	   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35H 9http://schemas.microsoft.com/winfx/2006/xaml/presentation      5   	   R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	     � Title$  Luban 表格管理工具��5   	   $	��600��6   $	��850��6+   .��5      � $��10q�6    � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $��* �6   �� 5	      $
��Auto �6   �� 5
      $	��120 �6   .+�5   
   ��  �� Orientation$ Vertical=�6   $��0,0,0,10q�6,   .+�5      �� $ 
Horizontal=�6   $��0,0,0,10q�62   .+�5      �� $��当前表格路径：��6   $��Center=�6+   $��0,0,5,0q�6F   �� 5      -   5         ��Name$ TxtPathDisplay��5      $	��450��64   $��Center=�6@   $��CharacterEllipsis=�6[   �� 5      -   5        $��选择路径��5      $	��100��6(   $��10,0,0,0q�64   6F   5      �� 5      $ 
Horizontal=�6   $��0,0,0,0q�62   .+�5      �� $��Luban导出脚本：��6   $��Center=�6.   $��0,0,5,0q�6I   �� 5      -   5        $ LubanBatPathTextBox��5      $	��450��69   $��Center=�6E   $��CharacterEllipsis=�6`   �� 5      -   5        $��选择脚本��5      $	��100��6(   64   5      5      �� 5!   
   -   5         ��Name$ ExcelListBox��5!      $��1��5"       �� SelectionMode$ Extended=�5#      $��0,0,0,10q�5$      $	��250��5%      5&      �� 5)   
   $��2��6   $ 
Horizontal=�6"   $
��Left=�6;   $��0,0,0,10q�6V   .+�5*      �� -   5        $��
SVN 更新��5*      $	��100��6&   $��0,0,10,0q�62   6D   �� 5+      -   5        $��
SVN 提交��5+      $	��100��6&   $��0,0,10,0q�62   6D   �� 5,      -   5        $��导出表格��5,      $	��150��6$   60   5-      �� 50   
   -	   5         ��Name$ LogBox��50      $��3��6"    �� 
IsReadOnly$
 True��51      " �� VerticalScrollBarVisibility$
 Auto=�52      	 �� TextWrapping$
	 Wrap=�53      $��Consolas)�54      ���#f4f4f455      56      57      