   M S B A M L   `   `   `  ����H  DDogsViewModel, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:DogsViewModelDogsViewModel  5   	   '    !DogsViewModel.OpenDirectoryWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   -localclr-namespace:DogsViewModel     5   	     � Title$  Open directory��5   	   $	��135��6    $	��480��6-   $	��135��69   $	��480��6I   $	��135��6X   $	��480��6h    � 
ResizeMode$ CanMinimize=�6w   .��5	      � ��5
   
   �� 5       �� EndPoint$ 0.5,1&�6"    �� 
StartPoint$ 0.5,0&�63   .��5      �  � Color$ Black��6    � Offset$ 0Q�6-   � 5      $ 	#FF309E53��6   $ 1Q�61   5       � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   �� -   5         ��Name$ Label1��5      $��0��6   $8��2Enter information about the directory being opened��6-   $
��Left=�6j   $��26��6�   $��Center=�6�   $	��306��6�   $
��Bold"�6�   
�������$��Arial)�6�   $��10,0,0,0q�6�   �� 5   
   -   5         ��Name$ TextBox1��5      $��1��6   $
��Left=�61   $��20��6L   	 �� TextWrapping$
	 Wrap=�6X   $"��Enter path to your directory��6l   $��Center=�6�   $	��377��6�   $��16,5,0,5q�6�   �� 5   
   -   5        
 ��Name$
 Button3��5      $��1��6   $��Browse��6/   $��Right=�6@   $��Center=�6\   $��76��6w   $��20��6�   $��0,5,5,5q�6�   �� 5   
   -   5        $
 Button1��5      $��2��6   $
��Open��6/   $��Right=�6>   $��Center=�6[   $��75��6v   $��
0,10,85,20q�6�   �� 5   
   -   5        $
 Button2��5      $��2��6   $��Cancel��6/   $��Right=�6@   $��Center=�6]   $��75��6x   $��	0,10,5,20q�6�   5      5      