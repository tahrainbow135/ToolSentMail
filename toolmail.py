# Cài python và cài thư viện
# Lệnh cài thư viện: pip install pandas openpyxl
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

GMAIL_USER = ""
GMAIL_PASSWORD = ""

# Đọc dữ liệu từ file Excel
# Đưa đường dẫn đến file Excel vào đây
df = pd.read_excel("D://User//Downloads//cv2.xlsx")

# Kết nối đến máy chủ Gmail
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(GMAIL_USER, GMAIL_PASSWORD)
successful_recipients = []
# Gửi email cho từng người
for index, row in df.iterrows():
    # Nhập email từ cột tên email
    email = row["Email"]
    # Nhập tên từ cột tên name
    name = row["Name"]

    # Nhập thời gian từ cột tên time (Ví dụ: 9h00, chủ nhật, ngày 24, tháng 09, năm 2023)
    time = row["Time"]
    # Nhập địa điểm từ cột tên location (Ví dụ: Nhà ăn Học viện Công nghệ Bưu chính Viễn thông) sao nó cứ ở đây???
    location = row["Location"]

    link_form = "https://forms.gle/CicSqQmWw6cL6dae7"

    # Tạo email
    msg = MIMEMultipart("alternative")
    msg["From"] = GMAIL_USER
    msg["To"] = email
    msg["Subject"] = "[CLB IT PTIT] THƯ MỜI SINH NHẬT 11 TUỔI"

    # Nội dung email thay name, time, location, link_form
    html_body = f"""
        <!DOCTYPE html>
        <html lang="en">
          <head>
            <meta charset="UTF-8" />
            <meta name="viewport" content="width=device-width, initial-scale=1.0" />
            <title>Document</title>
          </head>
          <body>
            <table
              id="m_-459558170298566609u_body"
              style="
                border-collapse: collapse;
                table-layout: fixed;
                border-spacing: 0;
                vertical-align: top;
                min-width: 320px;
                margin: 0 auto;
                background-color: #f9f9f9;
                width: 100%;
              "
              cellpadding="0"
              cellspacing="0"
            >
              <tbody>
                <tr style="vertical-align: top">
                  <td
                    style="
                      word-break: break-word;
                      border-collapse: collapse !important;
                      vertical-align: top;
                    "
                  >
                    <div style="padding: 0px; background-color: transparent">
                      <div
                        style="
                          margin: 0 auto;
                          min-width: 320px;
                          max-width: 600px;
                          word-wrap: break-word;
                          word-break: break-word;
                          background-color: transparent;
                        "
                      >
                        <div
                          style="
                            border-collapse: collapse;
                            display: table;
                            width: 100%;
                            height: 100%;
                            background-color: transparent;
                          "
                        >
                          <div
                            style="
                              max-width: 320px;
                              min-width: 600px;
                              display: table-cell;
                              vertical-align: top;
                            "
                          >
                            <div
                              style="
                                height: 100%;
                                width: 100% !important;
                                border-radius: 0px;
                              "
                            >
                              <div
                                style="
                                  box-sizing: border-box;
                                  height: 100%;
                                  padding: 0px;
                                  border-top: 5px solid #560707;
                                  border-left: 5px solid #560707;
                                  border-right: 5px solid #560707;
                                  border-bottom: 5px solid #560707;
                                  border-radius: 0px;
                                "
                              >
                                <table
                                  style="font-family: 'Cabin', sans-serif"
                                  role="presentation"
                                  cellpadding="0"
                                  cellspacing="0"
                                  width="100%"
                                  border="0"
                                >
                                  <tbody>
                                    <tr>
                                      <td
                                        style="
                                          word-break: break-word;
                                          padding: 0px;
                                          font-family: 'Cabin', sans-serif;
                                        "
                                        align="left"
                                      >
                                        <table
                                          width="100%"
                                          cellpadding="0"
                                          cellspacing="0"
                                          border="0"
                                        >
                                          <tbody>
                                            <tr>
                                              <td
                                                style="
                                                  padding-right: 0px;
                                                  padding-left: 0px;
                                                "
                                                align="center"
                                              >
                                                <img
                                                  align="center"
                                                  border="0"
                                                  src="https://ci3.googleusercontent.com/meips/ADKq_NZJXxtE5Lo6GAozFhUABP6AdfAiC9cjdBjHkhWSPpmBA1wR1mnjY1O_jiEW0RiugcK3xe8P3FMYO6blA9HXI3s8fIfzfVGPjg=s0-d-e1-ft#https://share1.cloudhq-mkt3.net/88385f26ce1da9.png"
                                                  alt=""
                                                  title=""
                                                  style="
                                                    outline: none;
                                                    text-decoration: none;
                                                    clear: both;
                                                    display: inline-block !important;
                                                    border: none;
                                                    height: auto;
                                                    float: none;
                                                    width: 100%;
                                                    max-width: 600px;
                                                  "
                                                  width="600"
                                                  class="CToWUd a6T"
                                                  data-bit="iit"
                                                  tabindex="0"
                                                />
                                              </td>
                                            </tr>
                                          </tbody>
                                        </table>
                                      </td>
                                    </tr>
                                  </tbody>
                                </table>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div style="padding: 0px; background-color: transparent">
                      <div
                        style="
                          margin: 0 auto;
                          min-width: 320px;
                          max-width: 600px;
                          word-wrap: break-word;
                          word-break: break-word;
                          background-color: #ffffff;
                        "
                      >
                        <div
                          style="
                            border-collapse: collapse;
                            display: table;
                            width: 100%;
                            height: 100%;
                            background-color: transparent;
                          "
                        >
                          <div
                            style="
                              max-width: 320px;
                              min-width: 600px;
                              display: table-cell;
                              vertical-align: top;
                            "
                          >
                            <div style="height: 100%; width: 100% !important">
                              <div
                                style="
                                  box-sizing: border-box;
                                  height: 100%;
                                  padding: 0px;
                                  border-top: 3px solid #560707;
                                  border-left: 3px solid #560707;
                                  border-right: 3px solid #560707;
                                  border-bottom: 3px solid #560707;
                                "
                              >
                                <table
                                  style="font-family: 'Cabin', sans-serif"
                                  role="presentation"
                                  cellpadding="0"
                                  cellspacing="0"
                                  width="100%"
                                  border="0"
                                >
                                  <tbody>
                                    <tr>
                                      <td
                                        style="
                                          word-break: break-word;
                                          padding: 20px 15px;
                                          font-family: 'Cabin', sans-serif;
                                        "
                                        align="left"
                                      >
                                        <div
                                          style="
                                            font-size: 14px;
                                            line-height: 160%;
                                            text-align: center;
                                            word-wrap: break-word;
                                          "
                                        >
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                              >Xin chào bạn {name},</span
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 10.5pt;
                                                font-family: Arial, sans-serif;
                                                color: #222222;
                                                line-height: 22.4px;
                                              "
                                              >&nbsp;</span
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                              >Lời đầu tiên,</span
                                            ><strong
                                              ><span
                                                style="
                                                  font-size: 13pt;
                                                  font-family: 'Times New Roman', serif;
                                                  color: #000000;
                                                  line-height: 27.2px;
                                                "
                                              >
                                              </span></strong
                                            ><strong
                                              ><span
                                                style="
                                                  font-size: 13pt;
                                                  font-family: 'Times New Roman', serif;
                                                  color: #ff9900;
                                                  line-height: 27.2px;
                                                "
                                                >IT PTIT</span
                                              ></strong
                                            ><span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                            >
                                              xin cảm ơn sự quan tâm của bạn dành cho sự
                                              kiện tuyển thành viên của Câu lạc bộ. Sau
                                              khi đọc hồ sơ của bạn, chúng tôi nhận thấy
                                              rằng bạn chính là mảnh ghép mà chúng tôi
                                              đang tìm kiếm.</span
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 10.5pt;
                                                font-family: Arial, sans-serif;
                                                color: #222222;
                                                line-height: 22.4px;
                                              "
                                              >&nbsp;</span
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                              >Chúng tôi xin</span
                                            ><strong
                                              ><span
                                                style="
                                                  font-size: 13pt;
                                                  font-family: 'Times New Roman', serif;
                                                  color: #ff9900;
                                                  line-height: 27.2px;
                                                "
                                              >
                                                Chúc mừng</span
                                              ></strong
                                            ><span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                            >
                                              bạn đã vượt qua vòng CV và lọt vào phỏng
                                              vấn của Câu lạc bộ. </span
                                            ><strong
                                              ><span
                                                style="
                                                  font-size: 13pt;
                                                  font-family: 'Times New Roman', serif;
                                                  color: #ff9900;
                                                  line-height: 27.2px;
                                                "
                                                >IT PTIT</span
                                              ></strong
                                            ><span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                            >
                                              hy vọng có thể trao đổi thêm với bạn trong
                                              buổi phỏng vấn này để hai bên có thể hiểu
                                              thêm về nhau hơn.&nbsp;</span
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 10.5pt;
                                                font-family: Arial, sans-serif;
                                                color: #222222;
                                                line-height: 22.4px;
                                              "
                                              >&nbsp;</span
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                              >Dưới đây là các thông tin chi tiết cho
                                              buổi phỏng vấn.&nbsp;</span
                                            >
                                          </p>
                                          <ul
                                            style="
                                              margin-top: 0px;
                                              margin-bottom: 0px;
                                              text-align: left;
                                            "
                                          >
                                            <li
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                font-weight: bold;
                                                margin-left: 11pt;
                                                line-height: 27.2px;
                                              "
                                            >
                                              <p
                                                style="
                                                  line-height: 160%;
                                                  background-color: rgb(255, 255, 255);
                                                  margin: 0px;
                                                "
                                              >
                                                <strong
                                                  ><span
                                                    style="
                                                      font-size: 13pt;
                                                      line-height: 27.2px;
                                                    "
                                                    >Thời gian: {time}</span
                                                  ></strong
                                                >
                                              </p>
                                            </li>
                                            <li
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                font-weight: bold;
                                                margin-left: 11pt;
                                                line-height: 27.2px;
                                              "
                                            >
                                              <p
                                                style="
                                                  line-height: 160%;
                                                  background-color: rgb(255, 255, 255);
                                                  margin: 0px;
                                                "
                                              >
                                                <strong
                                                  ><span
                                                    style="
                                                      font-size: 13pt;
                                                      line-height: 27.2px;
                                                    "
                                                    >Địa điểm: {location}</span
                                                  ></strong
                                                >
                                              </p>
                                            </li>
                                          </ul>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 10.5pt;
                                                font-family: Arial, sans-serif;
                                                color: #222222;
                                                line-height: 22.4px;
                                              "
                                              >&nbsp;</span
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                              >Ngoài ra, bạn có thể thay đổi thời gian
                                              phỏng vấn tại link sau đây: </span
                                            ><a
                                              href={link_form}
                                              style="text-decoration: none"
                                              target="_blank"
                                              data-saferedirecturl="https://www.google.com/url?q=https://forms.gle/fqhgX9wiKerosABZ9&amp;source=gmail&amp;ust=1725895552015000&amp;usg=AOvVaw3h_JpT87f34OLe4H-nfHuH"
                                              ><span
                                                style="
                                                  font-size: 13pt;
                                                  font-family: 'Times New Roman', serif;
                                                  color: #1155cc;
                                                  text-decoration: underline;
                                                  line-height: 27.2px;
                                                "
                                                >https://forms.gle/<wbr />fqhgX9wiKerosABZ9</span
                                              ></a
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 10.5pt;
                                                font-family: Arial, sans-serif;
                                                color: #222222;
                                                line-height: 22.4px;
                                              "
                                              >&nbsp;</span
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <strong
                                              ><span
                                                style="
                                                  font-size: 13pt;
                                                  font-family: 'Times New Roman', serif;
                                                  color: #000000;
                                                  line-height: 27.2px;
                                                "
                                                >Lưu ý:</span
                                              ></strong
                                            ><span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                            >
                                              Hãy đến sớm trước 10 phút để chuẩn bị tâm
                                              lý sẵn sàng và tự tin để có một buổi phỏng
                                              vấn thành công nhé. Chúng tôi mong rằng
                                              bạn sẽ có cơ hội thể hiện sự nhiệt huyết
                                              và kỹ năng của mình.&nbsp;</span
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            &nbsp;
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                              >Mọi thắc mắc hoặc câu hỏi khác, xin vui
                                              lòng gửi về Fanpage của Câu lạc bộ tại địa
                                              chỉ: </span
                                            ><a
                                              href="http://facebook.com/ITPTIT"
                                              style="text-decoration: none"
                                              target="_blank"
                                              data-saferedirecturl="https://www.google.com/url?q=http://facebook.com/ITPTIT&amp;source=gmail&amp;ust=1725895552015000&amp;usg=AOvVaw3l-1A7LaaHvLBhmYyOeBpk"
                                              ><span
                                                style="
                                                  font-size: 13pt;
                                                  font-family: 'Times New Roman', serif;
                                                  color: #000000;
                                                  text-decoration: underline;
                                                  line-height: 27.2px;
                                                "
                                                >http://facebook.com/ITPTIT</span
                                              ></a
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 10.5pt;
                                                font-family: Arial, sans-serif;
                                                color: #222222;
                                                line-height: 22.4px;
                                              "
                                              >&nbsp;</span
                                            >
                                          </p>
                                          <p
                                            style="
                                              line-height: 160%;
                                              background-color: rgb(255, 255, 255);
                                              text-align: left;
                                              margin: 0px;
                                            "
                                          >
                                            <span
                                              style="
                                                font-size: 13pt;
                                                font-family: 'Times New Roman', serif;
                                                color: #000000;
                                                line-height: 27.2px;
                                              "
                                              >Trân trọng và hẹn gặp bạn trong buổi
                                              phỏng vấn!</span
                                            >
                                          </p>
                                        </div>
                                      </td>
                                    </tr>
                                  </tbody>
                                </table>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div style="padding: 0px; background-color: transparent">
                      <div
                        style="
                          margin: 0 auto;
                          min-width: 320px;
                          max-width: 600px;
                          word-wrap: break-word;
                          word-break: break-word;
                          background-color: #560707;
                        "
                      >
                        <div
                          style="
                            border-collapse: collapse;
                            display: table;
                            width: 100%;
                            height: 100%;
                            background-color: transparent;
                          "
                        >
                          <div
                            style="
                              max-width: 320px;
                              min-width: 600px;
                              display: table-cell;
                              vertical-align: top;
                            "
                          >
                            <div style="height: 100%; width: 100% !important">
                              <div
                                style="
                                  box-sizing: border-box;
                                  height: 100%;
                                  padding: 0px;
                                  border-top: 0px solid transparent;
                                  border-left: 0px solid transparent;
                                  border-right: 0px solid transparent;
                                  border-bottom: 0px solid transparent;
                                "
                              >
                                <table
                                  style="font-family: 'Cabin', sans-serif"
                                  role="presentation"
                                  cellpadding="0"
                                  cellspacing="0"
                                  width="100%"
                                  border="0"
                                >
                                  <tbody>
                                    <tr>
                                      <td
                                        style="
                                          word-break: break-word;
                                          padding: 20px 10px;
                                          font-family: 'Cabin', sans-serif;
                                        "
                                        align="left"
                                      >
                                        <div align="center">
                                          <div style="display: table; max-width: 259px">
                                            <table
                                              align="center"
                                              border="0"
                                              cellspacing="0"
                                              cellpadding="0"
                                              width="32"
                                              height="32"
                                              style="
                                                width: 32px !important;
                                                height: 32px !important;
                                                display: inline-block;
                                                border-collapse: collapse;
                                                table-layout: fixed;
                                                border-spacing: 0;
                                                vertical-align: top;
                                                margin-right: 20px;
                                              "
                                            >
                                              <tbody>
                                                <tr style="vertical-align: top">
                                                  <td
                                                    align="center"
                                                    valign="middle"
                                                    style="
                                                      word-break: break-word;
                                                      border-collapse: collapse !important;
                                                      vertical-align: top;
                                                    "
                                                  >
                                                    <a
                                                      href="mailto:clb.it.ptit@gmail.com"
                                                      title="Email"
                                                      target="_blank"
                                                    >
                                                      <img
                                                        src="https://ci3.googleusercontent.com/meips/ADKq_NbQ1awrA278bNOI6nn55qKwaX22pUwD-08j8vmXPUc8AUsalrrDsPGcCRFSNd4w7-QcV9i4F1_J1NgzWdaLkuYzaxxBBpNQd3uVfe52XVpI6js=s0-d-e1-ft#https://cdn.tools.unlayer.com/social/icons/rounded/email.png"
                                                        alt="Email"
                                                        title="Email"
                                                        width="32"
                                                        style="
                                                          outline: none;
                                                          text-decoration: none;
                                                          clear: both;
                                                          display: block !important;
                                                          border: none;
                                                          height: auto;
                                                          float: none;
                                                          max-width: 32px !important;
                                                        "
                                                        class="CToWUd"
                                                        data-bit="iit"
                                                      />
                                                    </a>
                                                  </td>
                                                </tr>
                                              </tbody>
                                            </table>

                                            <table
                                              align="center"
                                              border="0"
                                              cellspacing="0"
                                              cellpadding="0"
                                              width="32"
                                              height="32"
                                              style="
                                                width: 32px !important;
                                                height: 32px !important;
                                                display: inline-block;
                                                border-collapse: collapse;
                                                table-layout: fixed;
                                                border-spacing: 0;
                                                vertical-align: top;
                                                margin-right: 20px;
                                              "
                                            >
                                              <tbody>
                                                <tr style="vertical-align: top">
                                                  <td
                                                    align="center"
                                                    valign="middle"
                                                    style="
                                                      word-break: break-word;
                                                      border-collapse: collapse !important;
                                                      vertical-align: top;
                                                    "
                                                  >
                                                    <a
                                                      href="https://www.facebook.com/ITPTIT"
                                                      title="Facebook"
                                                      target="_blank"
                                                      data-saferedirecturl="https://www.google.com/url?q=https://www.facebook.com/ITPTIT&amp;source=gmail&amp;ust=1725895552015000&amp;usg=AOvVaw07Wuag5xmF5An6fGz-r4eS"
                                                    >
                                                      <img
                                                        src="https://ci3.googleusercontent.com/meips/ADKq_NbuJJEY9VDc4xerFh35zfwU6rXm9N4x-QL2sA79wKkpfySrsmgkmKJ7Afkx1b-PqBBbzaqf1i0g7ldsxkRq56yaANUi_JXNkBa7T7HNWfS-l-Uey5A=s0-d-e1-ft#https://cdn.tools.unlayer.com/social/icons/rounded/facebook.png"
                                                        alt="Facebook"
                                                        title="Facebook"
                                                        width="32"
                                                        style="
                                                          outline: none;
                                                          text-decoration: none;
                                                          clear: both;
                                                          display: block !important;
                                                          border: none;
                                                          height: auto;
                                                          float: none;
                                                          max-width: 32px !important;
                                                        "
                                                        class="CToWUd"
                                                        data-bit="iit"
                                                      />
                                                    </a>
                                                  </td>
                                                </tr>
                                              </tbody>
                                            </table>

                                            <table
                                              align="center"
                                              border="0"
                                              cellspacing="0"
                                              cellpadding="0"
                                              width="32"
                                              height="32"
                                              style="
                                                width: 32px !important;
                                                height: 32px !important;
                                                display: inline-block;
                                                border-collapse: collapse;
                                                table-layout: fixed;
                                                border-spacing: 0;
                                                vertical-align: top;
                                                margin-right: 20px;
                                              "
                                            >
                                              <tbody>
                                                <tr style="vertical-align: top">
                                                  <td
                                                    align="center"
                                                    valign="middle"
                                                    style="
                                                      word-break: break-word;
                                                      border-collapse: collapse !important;
                                                      vertical-align: top;
                                                    "
                                                  >
                                                    <a
                                                      href="https://www.youtube.com/channel/UC8Iwsz8PT07_yVpqEvG7MRw"
                                                      title="YouTube"
                                                      target="_blank"
                                                      data-saferedirecturl="https://www.google.com/url?q=https://www.youtube.com/channel/UC8Iwsz8PT07_yVpqEvG7MRw&amp;source=gmail&amp;ust=1725895552015000&amp;usg=AOvVaw1XD8qZgtav45lZFv4fxogS"
                                                    >
                                                      <img
                                                        src="https://ci3.googleusercontent.com/meips/ADKq_NY06zLS_Qp0mU0LogYDcAPFvY3tHnEzKl1ZG2AzmOLekeQO-T6Qz-jdZKYYHyqVwflHbBZFHxNLIV8mQRErqSvYeTklqp7yTcaKa5N3AwZaQULSHg=s0-d-e1-ft#https://cdn.tools.unlayer.com/social/icons/rounded/youtube.png"
                                                        alt="YouTube"
                                                        title="YouTube"
                                                        width="32"
                                                        style="
                                                          outline: none;
                                                          text-decoration: none;
                                                          clear: both;
                                                          display: block !important;
                                                          border: none;
                                                          height: auto;
                                                          float: none;
                                                          max-width: 32px !important;
                                                        "
                                                        class="CToWUd"
                                                        data-bit="iit"
                                                      />
                                                    </a>
                                                  </td>
                                                </tr>
                                              </tbody>
                                            </table>

                                            <table
                                              align="center"
                                              border="0"
                                              cellspacing="0"
                                              cellpadding="0"
                                              width="32"
                                              height="32"
                                              style="
                                                width: 32px !important;
                                                height: 32px !important;
                                                display: inline-block;
                                                border-collapse: collapse;
                                                table-layout: fixed;
                                                border-spacing: 0;
                                                vertical-align: top;
                                                margin-right: 20px;
                                              "
                                            >
                                              <tbody>
                                                <tr style="vertical-align: top">
                                                  <td
                                                    align="center"
                                                    valign="middle"
                                                    style="
                                                      word-break: break-word;
                                                      border-collapse: collapse !important;
                                                      vertical-align: top;
                                                    "
                                                  >
                                                    <a
                                                      href="https://itptit.com/"
                                                      title="RSS"
                                                      target="_blank"
                                                      data-saferedirecturl="https://www.google.com/url?q=https://itptit.com/&amp;source=gmail&amp;ust=1725895552015000&amp;usg=AOvVaw2t7DO4XfpeRbHcz57QnGiT"
                                                    >
                                                      <img
                                                        src="https://ci3.googleusercontent.com/meips/ADKq_NZ3XYLwMpWWLMBQwGmDxnPNNqV1OvPr_PfwUQKj7w_j-Bi8G1uLXY7vzJGbM--3M491tauR0seRbCWthJDx9Agia5U8ht2ezRtBhjkIO5Wc=s0-d-e1-ft#https://cdn.tools.unlayer.com/social/icons/rounded/rss.png"
                                                        alt="RSS"
                                                        title="RSS"
                                                        width="32"
                                                        style="
                                                          outline: none;
                                                          text-decoration: none;
                                                          clear: both;
                                                          display: block !important;
                                                          border: none;
                                                          height: auto;
                                                          float: none;
                                                          max-width: 32px !important;
                                                        "
                                                        class="CToWUd"
                                                        data-bit="iit"
                                                      />
                                                    </a>
                                                  </td>
                                                </tr>
                                              </tbody>
                                            </table>

                                            <table
                                              align="center"
                                              border="0"
                                              cellspacing="0"
                                              cellpadding="0"
                                              width="32"
                                              height="32"
                                              style="
                                                width: 32px !important;
                                                height: 32px !important;
                                                display: inline-block;
                                                border-collapse: collapse;
                                                table-layout: fixed;
                                                border-spacing: 0;
                                                vertical-align: top;
                                                margin-right: 0px;
                                              "
                                            >
                                              <tbody>
                                                <tr style="vertical-align: top">
                                                  <td
                                                    align="center"
                                                    valign="middle"
                                                    style="
                                                      word-break: break-word;
                                                      border-collapse: collapse !important;
                                                      vertical-align: top;
                                                    "
                                                  >
                                                    <a
                                                      href="https://www.tiktok.com/@itclubptithn"
                                                      title="TikTok"
                                                      target="_blank"
                                                      data-saferedirecturl="https://www.google.com/url?q=https://www.tiktok.com/@itclubptithn&amp;source=gmail&amp;ust=1725895552015000&amp;usg=AOvVaw3ZFEljUOPjtKV-eXt8K9aE"
                                                    >
                                                      <img
                                                        src="https://ci3.googleusercontent.com/meips/ADKq_NaLjjxiDQ3q_x-6sJTxBD05lKZyuu4RnvlURDp4LnnIH8_-Rr7QjS76BOoJwsCMzMQ7U51QcQ1Gi8taGnmJ1y0L98-YYNVkKfqwd33YDDB_Yw0A=s0-d-e1-ft#https://cdn.tools.unlayer.com/social/icons/rounded/tiktok.png"
                                                        alt="TikTok"
                                                        title="TikTok"
                                                        width="32"
                                                        style="
                                                          outline: none;
                                                          text-decoration: none;
                                                          clear: both;
                                                          display: block !important;
                                                          border: none;
                                                          height: auto;
                                                          float: none;
                                                          max-width: 32px !important;
                                                        "
                                                        class="CToWUd"
                                                        data-bit="iit"
                                                      />
                                                    </a>
                                                  </td>
                                                </tr>
                                              </tbody>
                                            </table>
                                          </div>
                                        </div>
                                      </td>
                                    </tr>
                                  </tbody>
                                </table>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </body>
        </html>

    """

    msg.attach(MIMEText(html_body, "html"))

    # Gửi email
    server.sendmail(GMAIL_USER, email, msg.as_string())
    successful_recipients.append({"Name": name, "Email": email})

# Ngắt kết nối với máy chủ
server.quit()

print("Emails sent successfully!")
total_recipients = len(successful_recipients)

# Ghi danh sách người nhận và số lượng vào file
output_file = "D://User//Downloads//email_recipients.txt"

with open(output_file, "w", encoding="utf-8") as f:
    f.write(f"Tổng số người nhận email: {total_recipients}\n\n")
    f.write("Danh sách người nhận:\n")
    for recipient in successful_recipients:
        f.write(f"{recipient['Name']} - {recipient['Email']}\n")

print(f"Emails sent successfully! Total: {total_recipients}")