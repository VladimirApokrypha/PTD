   M S B A M L   `   `   `  ����H  DDogsViewModel, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:DogsViewModelDogsViewModel  5   	   &     DogsViewModel.NewDirectoryWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   -localclr-namespace:DogsViewModel     5   	   $	��155��5      $	��480��6'     � Title$  New directory��63   $	��155��6I   $	��480��6Y   $	��155��6h   $	��480��6x    � 
ResizeMode$ CanMinimize=�6�   .��5      � ��5	   
   �� 5
       �� EndPoint$ 0.5,1&�6"    �� 
StartPoint$ 0.5,0&�63   .��5      �  � Color$ Black��6    � Offset$ 0Q�6-   � 5      $ 	#FF309E53��6   $ 1Q�61   5       � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   �� -   5         ��Name$ TextBox2��5      $��2��6   $
��Left=�61   $��20��6L    �� TextWrapping$
 Wrap=�6X   $&�� Enter name of your new directory��6l   $��Center=�6�   $	��362��6�   $��16,6,0,5q�6�   6�   �� 5   
   -   5        	 ��Name$	 Button1��5      $��3��6   $��Create��6/   $��Right=�6@   $��Center=�6]   $��75��6x   $��0,0,85,0q�6�   �� 5   
   -   5        $	 Button2��5      $��3��6   $��Cancel��6/   $��Right=�6@   $��Center=�6]   $��75��6x   $��0,0,5,0q�6�   �� 5   
   -   5        $ TextBox1��5      $��1��6   $
��Left=�61   $��20��6L   $��16,0,0,5q�6X   $
 Wrap=�6j   $'��!Enter path for your new directory��6~   $	��Top=�6�   $	��368��6�   �� 5   
   -   5        $	 Button3��5      $��1��6   $��Browse��6/   $��Right=�6@   $��20��6\   $��0,0,5,5q�6h   $��Center=�6y   $��76��6�   6�   �� 5   
   -   5        
 ��Name$
 Label1��5      $��0��6   $0��*Enter information about your new directory��6-   $
��Left=�6b   $��27��6}   $��10,0,0,0q�6�   $��Center=�6�   $	��300��6�   $
��Bold"�6�   
�������$��Arial)�6�   5      5      