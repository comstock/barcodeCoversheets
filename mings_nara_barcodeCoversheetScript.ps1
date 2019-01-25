# Read a csv file, and print barcode coversheet
# the csv file have five columns, and has a header row
#  folder title,box number,folder number,start date,end date
#
# Note: 
# 1. using printhtml.exe (http://www.printhtml.com/)
# 2. using barcode generator from https://www.barcodesinc.com/generator

# some global vars. we can think to pass these as params if
# using this script for other projects.
$col = "Morgan v. Hennigan";
$cn = "99293814"

# making sure we get our csv filename
if ($args.Length -ne 1) {
    echo "A csv file is need";
    exit 1;
}
$fn = $args[0];

# let's just read the content of the csv file.
# we can also use import-csv command. but we feel we can not count on the header text to be
# consistent, since the file comes from someone else.  For example,
# $csv = Import-Csv $fn
# $csv | ForEach-Object { $_."Box Number"}
$fcontent = Get-Content $fn;

# starting from the second row, and print sheet for each row by calling printhtml.exe
for ($i=1; $i -lt $fcontent.Length; $i++) {
    # split the line into an array
    $parts=([string]$fcontent[$i]).split(',');
    # get the part we want
    $t = $parts[0].Trim();
    $bn = $parts[1].Trim();
    $fn = $parts[2].Trim();
    $sd = $parts[3].Trim();
    $ed = $parts[4].Trim();
    echo "print $t"
    # send the html content to printhtml to print the barcode page
    printhtml.exe html="<center><h1>$t, Box $bn, Folder $fn, $sd - $ed</h1><p/><p/><img border=0 src=https://www.barcodesinc.com/generator/image.php?code=$($cn)_$($bn)_$($fn)&style=196&type=C128B&width=1000&height=150&xres=3&font=5 /></center>" header="$col"
}
