   M S B A M L   `   `   `  ����H  DDogsViewModel, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:DogsViewModelDogsViewModel  5   	       DogsViewModel.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   -localclr-namespace:DogsViewModel     5   	     � Title$  Dogs Main Window ��5	   	   $	��340��6#   $	��575��60   $	��340��6<   $	��340��6L   $	��575��6\   $	��575��6k    � 
ResizeMode$ CanMinimize=�6z   .��5
      �  � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6    � ColumnDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   �� -   5         ��Name$ Button1��5      $��Generate report��6"   $��35��6<   $	��172��6H   $��3��6T   $��	0,10,0,10q�6a   $��Center=�6t   $��2��6�   $��Center=�6�   $��Center=�6�   $��Center=�6�   $��0��6  6  �� 5   
   -   5        $ Button2��5      $��Add contractor file��6"   $��Center=�6@   $��35��6]   $��Center=�6i   $	��172��6�   $��1��6�   $��2��6�   $��	0,60,0,10q�6�   $��0��6�   �� 5   
   -   5        $ Button3��5      $��Add subcontractor file��6"   $��Center=�6C   $��35��6`   $��Center=�6l   $	��172��6�   $��2��6�   $��2��6�   $��	0,10,0,10q�6�   $��0��6�   �� 5   
   -   5        $ 
Button1Vis��5      $��Generate report��6%   $��35��6?   $	��172��6K   $��3��6W   $��0,0,0,30q�6d   $��Center=�6v   $��Center=�6�   $��Center=�6�   $��Center=�6�   $��0��6�   $y�	Collapsed=�6  �� 5   
   -   5        $ 
Button2Vis��5      $��Add contractor file��6%   $��Center=�6C   $��35��6`   $��Center=�6l   $	��172��6�   $��1��6�   $��0,30,0,0q�6�   $��0��6�   $y�	Collapsed=�6�   �� 5   
   -   5        $ 
Button3Vis��5      $��Add subcontractor file��6%   $��Center=�6F   $��35��6c   $��Center=�6o   $	��172��6�   $��2��6�   $��0,0,0,0q�6�   $��0��6�   $y�	Collapsed=�6�   o� 5   
   $��20��6   $	��Top=�6   $��0��63   $��2��6@   $��Stretch=�6T   $	��575��6r   .@�5      m� $��Project��6    m� Icon 5      �� 5      $A��;DogsViewModel/DogsViewModel/Images/aryouqspvfpdlenizcfp.png��6   $��15��6a   $��15��6m   .@�5       m� -   5         m�Name$ NewProjectButton��5       $��New project��65   5!      m� 5"      -   5        $ OpenProjectButton��5"      $��Open project��66   5#      �� 5$      m� 5%      -	   5        $ 
ExitButton��5%      $
��Exit��6/   6Y   5&      m� 5'      $��Root directory��6    5(      �� 5)      $M��GDogsViewModel/DogsViewModel/Images/5b9c31c3ecf9e5afa33e837330dce277.png��6   $��12��6m   $��12��6y   .@�5+      m� -
   5        $ NewDirectoryButton��5+      $��New directory��67   6s   m� 5,      -   5        $ OpenDirectoryButton��5,      $��Open directory��68   6t   5-      m� 5.      -   5        $ ProjectTreeButton��5.      $��Project tree��62    5/      �� 50      $4��.DogsViewModel/DogsViewModel/Images/GetFile.png��6   $��12��6T   $��13��6`   52      m� 53      -   5        $ Marks��53      $��Marks��6&    54      �� 55      $7��1DogsViewModel/DogsViewModel/Images/img_356975.png��6   $��13��6W   $��12��6b   57      m� 58      -   5        $
 Help��58      $
��Help��6%    59      �� 5:      $7��1DogsViewModel/DogsViewModel/Images/img_452494.png��6   $��13��6W   $��12��6b   5<      5=      R� 5>   
   -   5         R�Name$ TreeOfProject��5>      $��1��6*   $	��200��67   $��Center=�6C   $	��Top=�6`   $	��250��6x   $��0,20,0,0q�6�   $��1��6�   $��3��6�   $y�	Collapsed=�6�   .@�5?      Q� $��Fruit��6   .@�5@      Q� $
��Lime��6   6/   Q� 5A      $��Orange��6   61   Q� 5B      $��Apple��6   60   5C      Q� 5D      $��Colors��6   .@�5E      Q� $	��Red��6   6.   Q� 5F      $��Orange��6   61   Q� 5G      $��Green��6   60   5H      5I      ��5J   
   �� 5K       �� EndPoint$ 0.5,1&�6"   	 � MappingMode$	 RelativeToBoundingBox=�63   
 �� 
StartPoint$
 0.5,0&�6W   .��5L      �  � Color$ Black��6   � 5M      $ 	#FF309E53��6    � Offset$ 1Q�61   5N      5P      5Q      