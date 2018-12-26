## Excel için VBA

Excel birçokları için çalışma hayatının önemli bir parçasıdır. Günlük işlerin büyük çoğunluğu Excel yardımı ile takip ediliyor, arşivleniyor, analiz ediliyor ya da raporlanıyor olabilir. Bu kadar sık kullanılsa da kullanıcının bilgi seviyesine göre aynı çalışma birkaç saniye, birkaç saat ya da birkaç hafta alabilir. 

Bu kursun amacı her seviyede ama özellikle başlangıç seviyesindeki kullanıcının Excel bilgisini artırarak, Excel ile takip ettiği iş süreçlerinde daha verimli olmasını sağlamaktır.

### İçerik

Bu kursun içeriği aşağıdaki gibidir:

- Vba Editörü ve Makro İçeren Dosyaların Özellikleri
- VBA Dili Yapısı ve Özellikleri
- Nesne Yapısı
- Olay Yapısı
- Hata Ayıklama
- Dış Kaynaklara Erişim ve Diğer Dillerle Entegrasyon

### Excel Dosyaları İçin Makro Ayarları

Excel içinde makro yazmak için Alt+F11 tuşları ile VBA editörüne erişilebilir. Ancak sık sık makro yazmak isteyecek bir kişinin Developer sekmesini aktif hale getirmesi gerekmektedir.
Developer sekmesi varsayılan durumda kapalıdır. Bu sekmeyi aktif hale getirmek için File->Options->Customize Ribbon yoluyla Customize the Ribbon alanına eriştikten sonra sayfanın sağ kısmında Customize the Ribbon alanında Developer seçeneğinin aktif hale getirilmesi gerekir.

![Activate Developer Tab](/img/activate_developer_tab.jpg)

Developer sekmesinde Code alanında Visual Basic düğmesi görülür. Bu düğmeye tıklayarak Visual Basic editörü açılır. Bunun dışında Macros düğmesi makrolar kutusunu açar. Recod Macro düğmesi makro kaydı başlatılmasını sağlar. Use Relative References makro kaydederken Range, Cells gibi alanların seçilen alanların tam anlamıyla kaydedilmesi yerine göreceli olarak kaydedilmesini sağlar. Mesela seçili hücrenin yazı tipini değiştirdiğinizde bu özellik aktif ise makro her çalıştırıldığında hangi hücre seçili ise o hücrenin yazı tipini değiştirecektir. Aksi taktirde seçilen hücreden bağımsız olarak tam anlamıyla kayıt yaparken hangi hücrenin yazı tipi değiştirildiyse, her çalıştırmada yine o hücrenin yazı tipi değiştirilecektir.
