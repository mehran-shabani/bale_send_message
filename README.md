# bale_send_message

Send fixed or personalized Bale Safir messages from Excel files.

## Excel format

Only a mobile-number column is required. A single-column workbook is accepted
with a header such as `موبایل`, `شماره`, `phone`, or `phone_number`, and is also
accepted without a header when its first cell is a valid Iranian mobile number.
Name and family-name columns are optional, and the message may be completely
static with no placeholders.

## Cost estimate

The documented Safir API does not expose a tariff or balance endpoint. Set the
current tariff for your account in `.env` to show a pre-send dashboard estimate:

```env
BALE_MESSAGE_PRICE_RIAL=0
```

Use the current value supplied in your Safir panel/contract. `0` disables the
estimate. The displayed total is an estimate based on valid, non-duplicate rows.
