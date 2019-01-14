# Excel için VBA

Excel çalışma hayatının önemli bir aracıdır. Günlük işlerin büyük çoğunluğu Excel yardımı ile takip ediliyor, arşivleniyor, analiz ediliyor ya da raporlanıyor olabilir. Excel bu kadar sık kullanılsa da, kullanıcının bilgi seviyesine göre aynı çalışma birkaç saniye, birkaç saat ya da birkaç hafta alabilir. 

Bu kursun amacı her seviyede ama özellikle başlangıç seviyesindeki kullanıcının Excel bilgisini artırarak, Excel ile takip ettiği iş süreçlerinde daha verimli olmasını sağlamaktır.

### İçerik

Bu dökümanın içeriği aşağıdaki gibidir:

- Vba Editörü ve Makro İçeren Dosyaların Özellikleri
- VBA Dili Yapısı ve Özellikleri
- Nesne Yapısı
- Olay Yapısı
- Hata Ayıklama
- Dış Kaynaklara Erişim ve Diğer Dillerle Entegrasyon

## VBA Editörü

Excel içinde makro yazmak için Alt+F11 tuşları ile VBA editörüne erişilebilir. Ancak sık sık makro yazmak isteyecek bir kişinin Developer sekmesini aktif hale getirmesi gerekmektedir.
Developer sekmesi varsayılan durumda kapalıdır. Bu sekmeyi aktif hale getirmek için File->Options->Customize Ribbon yoluyla Customize the Ribbon alanına eriştikten sonra sayfanın sağ kısmında Customize the Ribbon alanında Developer seçeneğinin aktif hale getirilmesi gerekir.

![Activate Developer Tab](/img/activate_developer_tab.jpg)

Developer sekmesinde Code alanında Visual Basic düğmesi (1) görülür. Bu düğmeye tıklayarak Visual Basic editörü açılır. 

![Developer Tab](/img/developer-tab.png)

Bunun dışında Macros düğmesi (2) makrolar kutusunu açar. Record Macro düğmesi (3) makro kaydı başlatılmasını sağlar. Use Relative References makro kaydederken Range, Cells gibi alanların seçilen alanların tam anlamıyla kaydedilmesi yerine göreceli olarak kaydedilmesini sağlar. Mesela seçili hücrenin yazı tipini değiştirdiğinizde bu özellik aktif ise makro her çalıştırıldığında hangi hücre seçili ise o hücrenin yazı tipini değiştirecektir. Aksi taktirde seçilen hücreden bağımsız olarak tam anlamıyla kayıt yaparken hangi hücrenin yazı tipi değiştirildiyse, her çalıştırmada yine o hücrenin yazı tipi değiştirilecektir.

### Hello World!

Visual Basic editörü açılarak makro yazmaya başlanabilir. Editör ilk kez açıldığında solda Project Explorer penceresi ve sağda boş alan yer alır. Project Explorer penceresinde açtığımız Excel dosyasının isminin de parantez içinde okunabileceği bir VBAProject yer alır. Uygulamada yüklü eklenti paketleri ya da açık diğer Excel dosyaları olması durumunda burada  birden fazla proje görülür.  

Makro kodu yazılmak istenilen proje üzerinde herhangi bir yere sağ tıklayarak Insert -> Module ile yeni bir modül oluşturulur. Yeni oluşturulan modüle çift tıklandığında ekranın boş olan sağ tarafında modül açılacaktır. Açılan pencerede kod yazılmaya başlanır.

Örnek olarak aşağıdaki kod yazılsın.

```VBA
Sub helloWorld()
  MsgBox "Hello World!"
End Sub
```

Bu kodu çalıştırmak için fare işaretçisi bu üç satırdan birinin üzerindeyken F5 tuşuna basılır veya Run menüsünden Run Sub/UserForm seçeneği seçilir veya standart araç çubuğundan yeşil üçgene tıklanır. 

Kod çalıştırıldığında Excel dosyasına geçilecek ve bir mesaj kutusu açılacaktır. Mesaj kutusunda OK seçeneği seçildiğinde yeniden editöre dönülür.

### Makro İçeren Excel Dosyaları

Makro içeren bir Excel dosyası normal olarak kaydedilirse kod içeren modül atılır ve tekrar kullanılamaz. Makro içeren bir Excel dosyasının (ve diğer ofis uygulamalarına ait dosyaların) makroyu kaybetmeden kaydedilmesi için Save As seçeneği seçilerek 'xlsm' uzantısı ile dosyanın kaydedilmesi gerekir. xlsm uzantısı makro içeren bir Excel dosyasının uzantısıdır. Bu dosyalar eğer farklı bir kaynaktan geliyorsa dikkatli olunması gerekir. Makro içeren bir dosya bilgi güvenliği risklerini de beraber getirir.

Microsoft bu nedenle varsayılan ayarlarda makro çalıştırılmasını engeller. Bu ayarlar Developer sekmesinde Macro Security tuşuna basılarak değiştirilebilir. Varsayılan olarak seçilen seçenek uyarı vermek kaydıyla tüm makroların engellendiği ikinci seçenektir. Bu seçenek önerilen seçenektir. Ancak bir makro geliştiricisi geçici olarak tüm makroların etkinleştirildiği dördüncü seçeneği de seçebilir. Tabi bu durumda tüm riskler kullanıcıya ait olacaktır.

### Veri Türleri

Aşağıdaki kodu dikkate alalım.

```VBA
Sub showSum()
  Dim a as Integer
  Dim b as Integer
  a = 1
  b = 2
  MsgBox a + b
End Sub
```

Kod çalıştırıldığında sonucu (3) görüntüleyen bir mesaj kutusu görülür. Bu kodda sırasıyla şunlar gerçekleşir:

1. a ve b ilk iki satırda tamsayı olarak tanımlanır. a ve b'ye **değişken** ismi verilir. dim değişken tanımlanırken kullanılan bir anahtar kelimedir. as değişkenin türü belirtilmeden önce kullanılan anahtar kelimedir. Integer değişkenlerin türünü (tamsayı) gösterir. 
2. Üçüncü ve dördüncü satırda değişkenlere sırasıyla 1 ve 2 değerleri atanır. = atama operatörü olarak isimlendirilir.
3. MsgBox ile mesaj kutusu çağırılır. Mesaj kutusunda a ve b değişkenlerinin toplamı gösterilir.

Şimdi toplamı göstermek için açıklayıcı bir ifade de eklensin.

```VBA
  ...
  MsgBox a & " ve " & b & "'nin toplamı " & a + b & "'dir."
  ...
```

Bu kez "1 ve 2'nin toplamı 3'tür." cümlesi görülür. Cümleyi bu kez bir değişkene atayarak bu değişken MsgBox'a aktarılsın.

```VBA
  ...
  Dim message As String
  message = a & " ve " & b & "'nin toplamı " & a + b & "'dir."
  MsgBox message
  ...
```

Yine aynı sonuç elde edilir. Burada yeni bir değişken türü kullanılmıştır. String değişkeni, değişkenin metin ifadeler aldığı durumlarda kullanılır. Metin ifadelerini birleştirmek için & birleştirme operatörü kullanılır. message ifadesine atanan metinde birleştirme operatörü kullanılmıştır. Birleştirme operatörü örnekte de göründüğü gibi metin ifadeleri dışında sayıları da birleştirmede kullanılabilir. Ancak elde edilen ifade bir her zaman metin olacaktır.

Sub ile başlayan ve End Sub ile biten ifade bir **prosedür** olarak isimlendirilir. Prosedürler paketlenmiş kod parçaları olarak düşünülebilir. Sıklıkla kullanılacak olan ve kendi içinde mantıklı bir bütünlüğe sahip olan kod satırları bir prosedür içinde birleştirilir.

Şimdi sıklıkla bahsi geçen değişken ve operatörlere daha yakından bakılsın.

Veri türleri

String ve Integer en sık kullanılan veri türlerindendir. Bu ikisi ve kullanılan diğer veri türlerinin bazıları aşağıdaki gibidir:

Veri Türü | Boyut | Aralık
----------|-------|-------
Byte | 1 Byte | 0 - 255
Boolean | 2 Byte | True ya da False
Integer | 2 Byte | -32,768 - 32,767
Long | 4 Byte | -2,147,483,648 - 2,147,483,647
LongLong | 8 Byte | -9,223,372,036,854,775,808 - 9,223,372,036,854,775,807
Single | 4 Byte | -3.402823E38 - -1.401298E-45 (negatif), 1.401298E-45 - 3.402823E38 (pozitif)
Double | 8 Byte | -1.79769313486231E308 - -4.94065645841247E-324 (negatif), 4.94065645841247E-324 - 1.79769313486232E308 (positif)
Date | 8 Byte | 1 Ocak 100 - 31 Aralık 9999
String | Metnin uzunluğu kadar | 1 - 65,400
Variant | 16 Byte | Double ya da String büyüklüğü kadar

Operatörler matematikte sıklıkla karşılaşılan ifadelerdir. + bir operatördür. Verilen iki sayının toplamını ifade eder. Operatörler verilen sınırlı sayıda girdiden çıktı üreten özel ifadelerdir. Yukarıdaki örnekte 1 ve 2 verildiğinde + operatörü 3 değerini üretir. VBA'da kullanılan diğer operatörler şöyledir.

Operatör Türü | Operatör | Açıklama
------------- | -------- | --------
Aritmetik operatörler | * | Verilen iki sayının çarpar
&#xfeff; | ^ | Birinci sayının ikinci sayı üssünü hesaplar
&#xfeff; | / | Birinci sayıyı ikinci sayıya böler (sonuç ondalıklı sayı)
&#xfeff; | \ | Birinci sayıyı ikinci sayıya böler (sonuç tamsayı)
&#xfeff; | Mod | Birinci sayının ikinci sayıya bölümünde kalanı verir
&#xfeff; | + | Verilen iki sayının toplar
&#xfeff; | - | Birinci sayıdan ikinci sayıyı çıkarır
Karşılaştırma operatörleri | = | Verilen iki sayının eşit olması durumunda True, aksi taktirde False değerini verir
&#xfeff; | Is | İki nesne aynı ise True aksi taktirde False değerini verir
&#xfeff; | Like | İki metin ifadesini karşılaştırarak yeterli eşleşme sağlandığında True diğer durumlarda False değerini verir
Birleştirme operatörleri | & | Verilen iki ifadeyi birleştirir
&#xfeff; | + | Verilen iki ifadeyi birleştirir
Mantıksal operatörler | And | Her iki ifade True ise True, diğer durumlarda False verir
&#xfeff; | Eqv | Her iki ifade True ya da her iki ifade False ise True diğer durumlarda False değerini verir
&#xfeff; | Imp | İlk ifade True ikinci ifade False ise False, diğer durumlarda True değerini verir
&#xfeff; | Not | İfade True ise False, False ise True değerini verir
&#xfeff; | Or | Her iki ifade False ise False diğer durumlarda True değerini verir
&#xfeff; | Xor | Her iki ifade True ya da her iki ifade False ise False diğer durumlarda True değerini verir

Mantıksal operatörlerde operandların (parametrelerin, ifadelerin...vb.) herhangi birisi ya da her ikisi de Null olduğunda True ve False dışında sonuç Null değeri olabilir. 

Bir önceki örneğe geri dönülürse, bu örnekte integer ve string veri türleri ile toplama, birleştirme ve atama operatörleri kullanılmıştır. Bu örneği hep aynı sayıların toplamını yazdırmaktansa kullanıcının seçeceği sayıların toplamını yazdıran bir prosedüre çevirmek daha yararlı olacaktır. Bu 2 şekilde yapılabilir.

Önce kullanıcıdan doğrudan sayıların istendiği versiyon incelensin.

```VBA
Sub showSum()
  Dim a as Integer, b as Integer
  Dim message As String
  a = InputBox("Toplanmasını istediğiniz ilk sayıyı giriniz")
  b = InputBox("Diğer sayıyı giriniz")
  message = a & " ve " & b & "'nin toplamı " & a + b & "'dir."
  MsgBox message
End Sub
```

InputBox ileti kutusu MsgBox gibi açılan bir kutudur. InputBox kullanıcının giriş yapabileceği bir girdi kutusuna sahiptir. InputBox ileti kutusu için kullanılabilecek üç parametre vardır. Bunlardan ilki prompt, ikincisi title, üçüncüsü ise defaulttur. Eğer parametre ismi girilmezse bu parametreler belirtilen sırada verilmelidir. Prompt kullanıcıya gösterilecek mesaj, title ise mesajın başlığıdır. Default parametresi varsayılan değer atanmasını sağlar. Varsayılan değer InputBox ile değeri kullanıcıdan istenen değişkene verilen ilk değerdir. Eğer kullanıcı giriş yapmazsa değişken bu değer ile işlem görmeye devam eder.

```VBA
  ...
  a = InputBox(Prompt:="Toplanmasını istediğiniz ilk sayıyı giriniz", Title:="İki sayının toplamı", Default:=5)
  ...
```

ya da kısaca

```VBA
  ...
  a = InputBox("Toplanmasını istediğiniz ilk sayıyı giriniz", "İki sayının toplamı", 5)
  ...
```

olarak kodlanabilir. Parametre ismi kullanıldığında parametrelerin hangi sırada kullanıldığının bir önemi kalmaz.

```VBA
  ...
  a = InputBox(Title:="İki sayının toplamı", Default:=5, Prompt:="Toplanmasını istediğiniz ilk sayıyı giriniz")
  ...
```

Parametrelerden sadece biri ya da sadece ikisi de belirtilebilir. Ancak Prompt zorunlu parametredir. Promptun her durumda belirtilmesi gerekir.

```VBA
  ...
  a = InputBox(Prompt:="Toplanmasını istediğiniz ilk sayıyı giriniz")
  ...
```

ya da

```VBA
  ...
  a = InputBox("Toplanmasını istediğiniz ilk sayıyı giriniz", "İki sayının toplamı")
  ...
```

Peki toplama işlemi için istenen değerler ileti kutusu ile değil de InputBox'da olduğu gibi parametreler ile istenseydi? Prosedür parametrelerle çağrıldığında parametrelerinin toplamını verecek şekilde de düzenlenebilir. 

```VBA
  Sub showSum(a As Integer, b As Integer)
    Dim message As String
    message = a & " ve " & b & "'nin toplamı " & a + b & "'dir."
    MsgBox message
  End Sub
```

Bu prosedür a ve b isminde iki parametreyi girdi olarak kabul eder. Bu iki parametrenin değeri toplanarak mesaj kutusunda görüntülenir. Ancak bu prosedür, editör penceresindeki Run düğmesi kullanılarak ya da Macros penceresinde seçilerek çalıştırılamaz. Macros penceresinde bu prosedür görünmeyecektir. Bu prosedür ancak başka bir prosedürden çağrılabilir.

```VBA
  Sub callSum()
    showSum 2, 4
  End Sub
```

Başka bir prosedürden çağrıldığında, çağrılan prosedürün ismi ve varsa parametreleri için kullanılacak değerler sırayla ya da isimleriyle, virgülle ayrılmış olarak girilir. Bu şekilde başka bir prosedürün içinden çağrılıyor olması gereksiz uzatılmış kod gibi görülebilir. Ancak bu kullanım kodu modüler hale getirmiştir. Bu sayede bu işleve yeniden ihtiyaç duyulduğunda tüm kodu tekrar yazmak yerine showSum prosedürü çağrılabilir.

```VBA
  Sub callSum()
    Call showSum(2, 4)
  End Sub
```

Prosedürler Call ile de çağrılabilir. Bu durumda parametreler parantez içinde girilmelidir.

showSum prosedürü şimdilik sadece sonucu bir mesaj kutusu içinde vermektedir. Eğer sonuç başka bir hesap içinde kullanılmak istenirse bir prosedür işe yaramaz. Bu durumda fonksiyon tanımlamak gerekir. Fonksiyonler bir çıktı üreten kod parçalarıdır. Çıktı olarak üretilen sonuç fonksiyon ismiyle aynı isimde bir değişkene atanmalıdır.

```VBA
  Function herzamanÜç() as Integer
    herzamanÜç = 3
  End Function
```

herzamanÜç fonksiyonu çağrıldığında 3 değerini verir. Bu fonksiyon herhangi bir girdiye ihtiyaç duymaz. showSum fonksiyonu şöyle tanımlanır.

```VBA
  Function showSum(a As Integer, b As Integer) As Integer
    showSum = a + b
  End Function
```

