## Excel için VBA

Excel çalışma hayatının önemli bir aracıdır. Günlük işlerin büyük çoğunluğu Excel yardımı ile takip ediliyor, arşivleniyor, analiz ediliyor ya da raporlanıyor olabilir. Excel bu kadar sık kullanılsa da, kullanıcının bilgi seviyesine göre aynı çalışma birkaç saniye, birkaç saat ya da birkaç hafta alabilir. 

Bu kursun amacı her seviyede ama özellikle başlangıç seviyesindeki kullanıcının Excel bilgisini artırarak, Excel ile takip ettiği iş süreçlerinde daha verimli olmasını sağlamaktır.

### İçerik

Bu kursun içeriği aşağıdaki gibidir:

- Vba Editörü ve Makro İçeren Dosyaların Özellikleri
- VBA Dili Yapısı ve Özellikleri
- Nesne Yapısı
- Olay Yapısı
- Hata Ayıklama
- Dış Kaynaklara Erişim ve Diğer Dillerle Entegrasyon

### Başlangıç Ayarları

Excel içinde makro yazmak için Alt+F11 tuşları ile VBA editörüne erişilebilir. Ancak sık sık makro yazmak isteyecek bir kişinin Developer sekmesini aktif hale getirmesi gerekmektedir.
Developer sekmesi varsayılan durumda kapalıdır. Bu sekmeyi aktif hale getirmek için File->Options->Customize Ribbon yoluyla Customize the Ribbon alanına eriştikten sonra sayfanın sağ kısmında Customize the Ribbon alanında Developer seçeneğinin aktif hale getirilmesi gerekir.

![Activate Developer Tab](/img/activate_developer_tab.jpg)

Developer sekmesinde Code alanında Visual Basic düğmesi görülür. Bu düğmeye tıklayarak Visual Basic editörü açılır. Bunun dışında Macros düğmesi makrolar kutusunu açar. Recod Macro düğmesi makro kaydı başlatılmasını sağlar. Use Relative References makro kaydederken Range, Cells gibi alanların seçilen alanların tam anlamıyla kaydedilmesi yerine göreceli olarak kaydedilmesini sağlar. Mesela seçili hücrenin yazı tipini değiştirdiğinizde bu özellik aktif ise makro her çalıştırıldığında hangi hücre seçili ise o hücrenin yazı tipini değiştirecektir. Aksi taktirde seçilen hücreden bağımsız olarak tam anlamıyla kayıt yaparken hangi hücrenin yazı tipi değiştirildiyse, her çalıştırmada yine o hücrenin yazı tipi değiştirilecektir.

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

1. a ve b ilk iki satırda tamsayı olarak tanımlanır. a ve b'ye değişken ismi verilir. dim değişken tanımlanırken kullanılan bir anahtar kelimedir. as değişkenin türü belirtilmeden önce kullanılan anahtar kelimedir. Integer değişkenlerin türünü (tamsayı) gösterir. 

2. Üçüncü ve dördüncü satırda değişkenlere sırasıyla 1 ve 2 değerleri atanır. = atama operatörü olarak isimlendirilir.

3. MsgBox ile mesaj kutusu çağırılır. Mesaj kutusunda a ve b değişkenlerinin toplamı gösterilir.

Şimdi toplamı göstermek için açıklayıcı bir ifade de eklensin.

```VBA
  MsgBox a & " ve " & b & "'nin toplamı " & a + b & "'tür."
```

Bu kez "1 ve 2'nin toplamı 3'tür." cümlesi görülür. Cümleyi bu kez bir değişkene atayarak bu değişken MsgBox'a aktarılsın.

```VBA
  Dim message As String
  message = a & " ve " & b & "'nin toplamı " & a + b & "'tür."
  MsgBox message
```

Yine aynı sonuç elde edilir. Burada yeni bir değişken türü kullanılmıştır. String değişkeni, değişkenin metin ifadeler aldığı durumlarda kullanılır. Metin ifadelerini birleştirmek için & birleştirme operatörü kullanılır. message ifadesine atanan metinde birleştirme operatörü kullanılmıştır. Birleştirme operatörü örnekte de göründüğü gibi metin ifadeleri dışında sayıları da birleştirmede kullanılabilir. Ancak elde edilen ifade bir her zaman metin olacaktır.

Şimdi sıklıkla bahsi geçen değişken ve operatörlere daha yakından bakılsın.

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

