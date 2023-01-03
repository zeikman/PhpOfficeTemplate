<?php
/*
Copyright 2007-2008 Brenton Fletcher. http://bloople.net/num2text
You can use this freely and modify it however you want.
*/
function convertNumberIndo($num)
{
   list($num, $dec) = explode(".", $num);

   $output = "";

   if($num{0} == "-")
   {
      $output = "negatif ";
      $num = ltrim($num, "-");
   }
   else if($num{0} == "+")
   {
      $output = "positif ";
      $num = ltrim($num, "+");
   }

   if($num{0} == "0")
   {
      $output .= "kosong";
   }
   else
   {
      $num = str_pad($num, 36, "0", STR_PAD_LEFT);
      $group = rtrim(chunk_split($num, 3, " "), " ");
      $groups = explode(" ", $group);

      $groups2 = array();
      foreach($groups as $g) $groups2[] = convertThreeDigitIndo($g{0}, $g{1}, $g{2});

      for($z = 0; $z < count($groups2); $z++)
      {
         if($groups2[$z] != "")
         {
            $output .= $groups2[$z].convertGroupIndo(11 - $z).($z < 11 && !array_search('', array_slice($groups2, $z + 1, -1))
             && $groups2[11] != '' && $groups[11]{0} == '0' ? " dan " : " ");
         }
      }

      $output = rtrim($output, ", ");
   }

   if($dec > 0)
   {
      $output .= " perpuluhan";
      for($i = 0; $i < strlen($dec); $i++) $output .= " ".convertDigitIndo($dec{$i});
   }

   return $output;
}

function convertGroupIndo($index)
{
   switch($index)
   {
      case 11: return " decillion";
      case 10: return " nonillion";
      case 9: return " octillion";
      case 8: return " septillion";
      case 7: return " sextillion";
      case 6: return " quintrillion";
      case 5: return " quadrillion";
      case 4: return " trilion";
      case 3: return " bilion";
      case 2: return " juta";
      case 1: return " ribu";
      case 0: return "";
   }
}

function convertThreeDigitIndo($dig1, $dig2, $dig3)
{
   $output = "";

   if($dig1 == "0" && $dig2 == "0" && $dig3 == "0") return "";

   if($dig1 != "0")
   {
      $output .= convertDigitIndo($dig1)." ratus";
      if($dig2 != "0" || $dig3 != "0") $output .= " dan ";
   }

   if($dig2 != "0") $output .= convertTwoDigitIndo($dig2, $dig3);
   else if($dig3 != "0") $output .= convertDigitIndo($dig3);

   return $output;
}

function convertTwoDigitIndo($dig1, $dig2)
{
   if($dig2 == "0")
   {
      switch($dig1)
      {
         case "1": return "sepuluh";
         case "2": return "dua puluh";
         case "3": return "tiga puluh";
         case "4": return "empat puluh";
         case "5": return "lima puluh";
         case "6": return "enam puluh";
         case "7": return "tujuh puluh";
         case "8": return "lapan puluh";
         case "9": return "sembilan puluh";
      }
   }
   else if($dig1 == "1")
   {
      switch($dig2)
      {
         case "1": return "sebelas";
         case "2": return "dua belas";
         case "3": return "tiga belas";
         case "4": return "empat belas";
         case "5": return "lima belas";
         case "6": return "enam belas";
         case "7": return "tujuh belas";
         case "8": return "lapan belas";
         case "9": return "sembilan belas";
      }
   }
   else
   {
      $temp = convertDigitIndo($dig2);
      switch($dig1)
      {
         case "2": return "dua puluh $temp";
         case "3": return "tiga puluh $temp";
         case "4": return "empat puluh $temp";
         case "5": return "lima puluh $temp";
         case "6": return "enam puluh $temp";
         case "7": return "tujuh puluh $temp";
         case "8": return "lapan puluh $temp";
         case "9": return "sembilan puluh $temp";
      }
   }
}

function convertDigitIndo($digit){
   switch($digit)   {
      case "0": return "kosong";
      case "1": return "satu";
      case "2": return "dua";
      case "3": return "tiga";
      case "4": return "empat";
      case "5": return "lima";
      case "6": return "enam";
      case "7": return "tujuh";
      case "8": return "lapan";
      case "9": return "sembilan";
   }
}
?>
