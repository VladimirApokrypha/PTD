   M S B A M L   `   `   `  ����H  DDogsViewModel, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:DogsViewModelDogsViewModel  5   	   %    DogsViewModel.OpenProjectWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   -localclr-namespace:DogsViewModel     5   	     � Title$  Open project��5   	   $	��135��6   $	��480��6+   $	��135��67   $	��480��6G   $	��135��6V   $	��480��6f    � 
ResizeMode$ CanMinimize=�6u   .��5	      � ��5
   
   �� 5       �� EndPoint$ 0.5,1&�6"    �� 
StartPoint$ 0.5,0&�63   .��5      �  � Color$ Black��6    � Offset$ 0Q�6-   � 5      $ 	#FF309E53��6   $ 1Q�61   5       � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   �� -   5         ��Name$ Label1��5      $��0��6   $6��0Enter information about the project being opened��6-   $
��Left=�6h   $��26��6�   $��10,0,0,0q�6�   $��Center=�6�   $	��294��6�   $
��Bold"�6�   
�������$��Arial)�6�   �� 5   
   -   5         ��Name$ TextBox1��5      $��1��6   $
��Left=�61   $��20��6L   $��16,5,0,5q�6X   	 �� TextWrapping$
	 Wrap=�6j   $ ��Enter path to your project��6~   $��Center=�6�   $	��377��6�   �� 5   
   -   5        
 ��Name$
 Button1��5      $��2��6   $
��Open��6/   $��Right=�6>   $��
0,10,85,20q�6Z   $��Center=�6n   $��74��6�   �� 5   
   -   5        $
 Button2��5      $��2��6   $��Cancel��6/   $��Right=�6@   $��	0,10,5,20q�6\   $��Center=�6o   $��75��6�   6�   �� 5   
   -   5        $
 Button3��5      $��1��6   $��Browse��6/   $��Right=�6@   $��0,5,5,5q�6\   $��Center=�6m   $��76��6�   $��20��6�   5      5      