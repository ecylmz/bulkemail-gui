# encoding: utf-8

require "gtk2"
require "roo"
require "tlsmail"
require "time"

class MailerGtk
  def initialize
    @window = Gtk::Window.new
    @window.set_title "Bulk Email Sender"
    @window.set_size_request(900,500)
    @window.set_resizable(true)
    @window.signal_connect('destroy'){Gtk.main_quit}
    @window.set_window_position(Gtk::Window::POS_CENTER)
    @editor = Gtk::TextView.new
    @subject = Gtk::Entry.new
    @from = Gtk::Entry.new
    @password = Gtk::Entry.new
    @password.visibility = false
    window_contain
    @window.show_all
  end

  def window_contain
    @window.set_border_width 15

    table = Gtk::Table.new 15, 15, false
    table.set_column_spacings 3

    from_title = Gtk::Label.new "Kimden"
    halign2 = Gtk::Alignment.new 0, 0, 0, 0
    halign2.add from_title
    table.attach(halign2, 0, 1, 0, 1, Gtk::FILL, Gtk::FILL, 0, 0)
    table.attach(@from, 0, 1, 1, 2, Gtk::FILL, Gtk::FILL, 0, 0)

    password_title = Gtk::Label.new "Parola"
    halign3 = Gtk::Alignment.new 0, 0, 0, 0
    halign3.add password_title
    table.attach(halign3, 0, 1, 2, 3, Gtk::FILL, Gtk::FILL, 0, 0)
    table.attach(@password, 0, 1, 3, 4, Gtk::FILL, Gtk::FILL, 0, 0)

    subject_title = Gtk::Label.new "Konu"
    halign = Gtk::Alignment.new 0, 0, 0, 0
    halign.add subject_title
    table.attach(halign, 0, 1, 4, 5, Gtk::FILL, Gtk::FILL, 0, 0)
    table.attach(@subject, 0, 1, 5, 6, Gtk::FILL, Gtk::FILL, 0, 0)

    email_title = Gtk::Label.new "Email"
    halign1 = Gtk::Alignment.new 0, 0, 0, 0
    halign1.add email_title
    table.attach(halign1, 0, 1, 6, 7, Gtk::FILL, Gtk::FILL, 0, 0)
    @editor.set_size_request 700, 300
    table.attach(@editor, 0, 1, 7, 8, Gtk::FILL, Gtk::FILL, 0, 0)

    activate = Gtk::Button.new "Liste Seç"
    activate.set_size_request 80, 30
    table.attach(activate, 3, 4, 1, 2, Gtk::FILL, Gtk::SHRINK, 1, 1)
    activate.signal_connect "clicked" do
      file_chooser
    end

    activate = Gtk::Button.new "Ek Seç"
    activate.set_size_request 80, 30
    table.attach(activate, 4, 5, 1, 2, Gtk::FILL, Gtk::SHRINK, 1, 1)
    activate.signal_connect "clicked" do
      attachment_chooser
    end

    valign = Gtk::Alignment.new 0, 0, 0, 0
    send = Gtk::Button.new "Gönder"
    send.set_size_request 80, 30
    send.signal_connect "clicked" do
      send_email
    end
    valign.add send
    table.set_row_spacing 1, 3
    table.attach(valign, 3, 4, 2, 3, Gtk::FILL, Gtk::FILL | Gtk::EXPAND, 1, 1)

    @window.add table
  end

  def file_chooser
    file_choose = Gtk::FileChooserDialog.new("Liste Seç",
                                              @window,
                                              Gtk::FileChooser::ACTION_OPEN,
                                              nil,
                                              [Gtk::Stock::CANCEL, Gtk::Dialog::RESPONSE_CANCEL],
                                              [Gtk::Stock::OPEN, Gtk::Dialog::RESPONSE_ACCEPT])

    if file_choose.run == Gtk::Dialog::RESPONSE_ACCEPT
      @filename = file_choose.filename
    end
    file_choose.destroy
  end

  def attachment_chooser
    file_choose = Gtk::FileChooserDialog.new("Ek Seç",
                                              @window,
                                              Gtk::FileChooser::ACTION_OPEN,
                                              nil,
                                              [Gtk::Stock::CANCEL, Gtk::Dialog::RESPONSE_CANCEL],
                                              [Gtk::Stock::OPEN, Gtk::Dialog::RESPONSE_ACCEPT])

    if file_choose.run == Gtk::Dialog::RESPONSE_ACCEPT
      @attachment = file_choose.filename
    end
    file_choose.destroy
  end

  def email_list
    list = []
    if @filename
      begin
        xls = Roo::Spreadsheet.open(@filename)
      rescue
        md = Gtk::MessageDialog.new(@window,
                                    Gtk::Dialog::DESTROY_WITH_PARENT, Gtk::MessageDialog::WARNING,
                                    Gtk::MessageDialog::BUTTONS_CLOSE, "Hatalı dosya formatı!")
        md.run
        md.destroy
        return nil
      end

      # excel dosyasındaki ilk sütunu al
      xls.each do |column|
      list << column[0]
      end
    else
      md = Gtk::MessageDialog.new(@window,
                                  Gtk::Dialog::DESTROY_WITH_PARENT, Gtk::MessageDialog::WARNING,
                                  Gtk::MessageDialog::BUTTONS_CLOSE, "Emaillerin ilk sütunda olduğu excel dosyanızı seçin!")
      md.run
      md.destroy
      return nil
    end
    return list
  end

  def send_email

    # kullanıcı adı ve parola boş mu kontrolü
    if @from.text.empty? or @password.text.empty?
      md = Gtk::MessageDialog.new(@window,
                                  Gtk::Dialog::DESTROY_WITH_PARENT, Gtk::MessageDialog::WARNING,
                                  Gtk::MessageDialog::BUTTONS_CLOSE, "Kimden ve Parola Kısmı Boş Bırakılamaz!")
      md.run
      md.destroy
      return
    end


    # smtp settings
    email_domain = @from.text.split("@").last
    if email_domain == "fesss.org"
      smtp_domain = "mail.fesss.org"
      domain = "fesss.org"
      port = 587
    elsif email_domain == "gmail.com"
      smtp_domain = "smtp.gmail.com"
      domain = "gmail.com"
      port = 587
    end


    if email_list
      Net::SMTP.enable_tls(OpenSSL::SSL::VERIFY_NONE)
      Net::SMTP.start(smtp_domain, port, domain, @from.text, @password.text, :login) do |smtp|
        md = Gtk::MessageDialog.new(@window,
                                    Gtk::Dialog::DESTROY_WITH_PARENT, Gtk::MessageDialog::WARNING,
                                    Gtk::MessageDialog::BUTTONS_CLOSE, "Bu Pencereyi Kapattıktan Sonra Gönderim Başlayacaktır!
                                    Mail Gönderimi Bittikten Sonra Ekrana Bilgilendirme Penceresi Gelecektir. Lütfen Bekleyiniz!")
        md.run
        md.destroy

        email_list.each do |to|
          begin
            # mail içeriğini hazırla
            if @attachment
              filename = File.basename(@attachment)
              file = open(@attachment, "rb") {|io| io.read}
              encode_file = [file].pack("m")
              marker = "AUNIQUEMARKER"
              content = <<EOF
From: #{@from.text}
To: #{to}
Subject: #{@subject.text}
Date: #{Time.now.rfc2822}
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary=#{marker}
--#{marker}
Content-Type: text/plain
Content-Transfer-Encoding:8bit

#{@editor.buffer.text}
--#{marker}
Content-Type: multipart/mixed; name=\"#{filename}\"
Content-Transfer-Encoding:base64
Content-Disposition: attachment; filename="#{filename}"

#{encode_file}
##--#{marker}--
EOF
            else
              content = <<EOF
From: #{@from.text}
To: #{to}
Subject: #{@subject.text}
Date: #{Time.now.rfc2822}

#{@editor.buffer.text}
EOF
            end
            # email'i gönder
            smtp.send_message(content, @from.text, to)
          rescue
            md = Gtk::MessageDialog.new(@window,
                                        Gtk::Dialog::DESTROY_WITH_PARENT, Gtk::MessageDialog::ERROR,
                                        Gtk::MessageDialog::BUTTONS_CLOSE, "Mail gönderilirken hata!\n
                                                                          Kullanıcı bilgilerinizi ve internet bağlantınızı kontrol edin!")
            md.run
            md.destroy
          end
        end
        md = Gtk::MessageDialog.new(@window,
                                    Gtk::Dialog::DESTROY_WITH_PARENT, Gtk::MessageDialog::INFO,
                                    Gtk::MessageDialog::BUTTONS_CLOSE, "Email Tüm Listeye Gönderildi!")
        md.run
        md.destroy
      end
    end
  end
end

app = MailerGtk.new
Gtk.main
