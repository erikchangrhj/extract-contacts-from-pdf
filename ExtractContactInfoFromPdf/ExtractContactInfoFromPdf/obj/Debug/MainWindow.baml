   M S B A M L   `   `   `  ����T  PExtractContactInfoFromPdf, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullE'clr-namespace:ExtractContactInfoFromPdfExtractContactInfoFromPdf  5   	   *    $ExtractContactInfoFromPdf.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   9local'clr-namespace:ExtractContactInfoFromPdf     5   	   -   5          � Title$  Convert PDF To CSV��5   	   $	��700��6$   $	��600��61    � Topmost .  � WindowStartupLocation$ CenterScreen=�6m    � Icon$ icon_excel.ico��6�   .��5	      � .+�5
   
   �� -   5         ��Name$ 
openButton��5
      $��Open PDF��6%   $
��Left=�68   $��	10,10,0,0q�6S   $	��Top=�6f   $	��100��6~   $��30��6�   �� 5   
   $��From:��6   $
��Left=�6    $��	10,55,0,0q�6;   $	��Top=�6N   $��30��6f   $��45��6r   $��14(�6}   �� 5   
   -   5         ��Name$ startPageNumberTextBox��5      $
��Left=�62   $��30��6M   $��	60,55,0,0q�6Y    �� TextWrapping$
 Wrap=�6l   $	��Top=�6�   $��50��6�   $��14(�6�   $��1��6�   �� 5   
   $��From:��6   $
��Left=�6    $��	10,95,0,0q�6;   $	��Top=�6N   $��30��6f   $��45��6r   $��14(�6}   �� 5   
   -   5        $ pageNumbersTextBox��5      $
��Left=�6.   $��30��6I   $��	60,95,0,0q�6U   $
 Wrap=�6h   $	��Top=�6|   $��50��6�   $��14(�6�   $
��1000��6�   �� 5   
   -   5        $ textBox��5      $
��Left=�6#   $	��649��6>   $��
125,10,0,0q�6K   $
 Wrap=�6_   $	��Top=�6s   $	��457��6�   $��14(�6�   5      5      