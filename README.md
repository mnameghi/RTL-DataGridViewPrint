RTL DataGridViewPrint provides simple right-to-left capability for printing a windows form datagridview content.By this way reporting is fast and simple. Your report printed with PrintDocument component.


## Quick start

To get started, Add DataGridViewPrint.cs class to your project. On a form with DataGridView add a PrintDocument component and use follow code:

```c#
DataGridViewPrint dg = new DataGridViewPrint(dataGridView1, printDocument1);
dg.PageTitle.HeaderStr = "Title for report";
dg.PageTitle.SubTitle1 = "subtitle 1";
dg.PageTitle.SubTitle2 = "subtitle 2";
dg.PageTitle.TitleFont = new System.Drawing.Font("B Mitra", 12f);
dg.WrapText = false;
dg.PrinPreview();
```

## Development

Contributors to RTL-DataGridViewPrint must agree the license by signing on the bottom of the `CONTRIBUTORS.md` file. To contribute:

- [fork the bootstrap-rtl repository](https://github.com/mnameghi/RTL-DataGridViewPrint/fork).
- make your changes
- *first time contributors*: Sign `CONTRIBUTORS.md` by adding your github username, full name, email address, and first contribution date. As follows:
    `YYYY/MM/DD, Github Username, Full Name, Email Address`
- commit your changes.
- send a pull request.


***Important Notice:* Before submitting any pull request, please ensure that your changes merge cleanly with the main repository at mnameghi/RTL-DataGridViewPrint.**


### Feature requests, and bug fixes

If you want a feature or a bug fixed, [report it via project's issues tracker](https://github.com/mnameghi/RTL-DataGridViewPrint/issues). However, if it's something you can fix yourself, *fork* the project, *do* whatever it takes to resolve it, and finally submit a *pull* request. I will personally thank you, and add your name to the list of contributors.

## Author

**Mahdi Nameghi**

+ [http://github.com/mnameghi](http://github.com/mnameghi)


## Copyright and license

This code released under [the MIT license](LICENSE). Docs released under [Creative Commons](docs/LICENSE).


