msg = EmailMessage()
msg['Subject'] = 'Missing tracking number'
msg['From'] = email_address
msg['To'] = ', '.join(recipients)
msg.set_content(f'No tracking number found in email {i} \nwhere customer shipping address is: {full_address} \nand order number is: {order} \ncopy and paste this link in browser to track order {href} \n\n\nP.S. There might be more issues with this email.')
msg.add_alternative(f'''
    <html>
        <p>
            No tracking number found in email {i}<br /> 
            where customer shipping address is: {full_address}<br />
            and the order number is {order}<br /><br /><br />
            P.S. There might be more issues with this email
        </p>
        <a href="{href}">Track Order</a>
    </html>    
        ''', subtype='html')
smtp.send_message(msg=msg)