using System;
using TFlex;
using TFlex.Model;
using TFlex.Command;
using System.Collections.Generic;
using System.Linq;
using TFlex.Configuration;
using TFlex.Model.Data.ProductStructure;
using Microsoft.SqlServer.Server;

namespace TFlexCADCommandsUser
{
    public class Factory : PluginFactory
    {
        public override Plugin CreateInstance()
        {
            return new TFlexCADCommandsUser(this);
        }
        public override Guid ID
        {
            get { return new Guid("{0C84154A-0469-47D1-B930-3F3D098E0F72}"); }
        }
        public override string Name
        {
            get { return "Formats_TheSecondSolution"; }
        }
    }
    enum Commands
    {
        myCommand1 = 1,
        myCommand2,
        myCommand3,
        myCommand4,
        myCommand5,
        myCommand6,
        myCommand7,
        myCommand8,
        myCommand9,
        myCommand10,
        myCommand11,
    };//Список команд

    class TFlexCADCommandsUser : Plugin
    {
        public TFlexCADCommandsUser(Factory factoty) : base(factoty)
        {
        }
        System.Drawing.Icon LoadIconResource(string name)
        {
            System.IO.Stream stream = GetType().Assembly.GetManifestResourceStream("TFlexCADCommandsUser.Icon_Files." + name + ".ico");
            return new System.Drawing.Icon(stream);
        }
        protected override void OnInitialize()
        {
            base.OnInitialize();
        }
        private RibbonGroup _group;
        private List<int> _listButton = new List<int>();
        private string[] TEXTs = new string[11] { "А4", "А3", "А2", "А1", "А0", "А4х3", "А3х3", "А2х3", "А1х3", "А0х2", "БЧ" };
        protected override void OnCreateTools()//Создание команд и регистрация кнопок в ленте
        {
            string nameGroup = "Форматы";
            for (int i = 1; i < 23; i++)
            {
                if (i < 12)
                {
                    RegisterCommand(i, TEXTs[i - 1], LoadIconResource(TEXTs[i - 1]), LoadIconResource(TEXTs[i - 1]));
                }
                else
                {
                    RegisterCommand(i, TEXTs[(i - 1) % 11], LoadIconResource(TEXTs[(i - 1) % 11] + "В"), LoadIconResource(TEXTs[(i - 1) % 11] + "В"));
                }
            }

            RibbonTab ribbonTab = new RibbonTab(nameGroup, this.ID, this);
            _group = ribbonTab.AddGroup(nameGroup);

            for (int i = 1; i < 12; i++)
            {
                _listButton.Add(i);
                _group.AddButton(i, TEXTs[i - 1], this, RibbonButtonStyle.LargeIconAndCaption);
            }

        }

        private void Replacing(int id, bool flag=false)//Замена иконок в ленте (вертикальное/горизонтальное положение)
        {
            if (flag)
            {
                var list = _group.Buttons.ToList();
                foreach (var button in _group.Buttons) { button.Remove(); }

                for (int i = id < 12 ? 1 : 12; i < (id < 12 ? 12 : 23); ++i)
                {
                    if (_listButton[(i - 1) % 11] == id) _listButton[(i - 1) % 11] += id >= 12 ? -11 : 11;
                    _group.AddButton(_listButton[(i - 1) % 11], TEXTs[(i - 1) % 11], this, RibbonButtonStyle.LargeIconAndCaption);
                }
            }
            else
            {
                var document = TFlex.Application.ActiveDocument;
                var activPage = document.ActivePage;
                document.BeginChanges("Изменить формат");
                activPage.Properties.Paper.Landscape = id >= 12 ? LandscapePaper.Vertical : LandscapePaper.Horizontal;
                document.EndChanges();
            }
        }
        protected override void PluginCommandEventHandler(PluginCommandEventArgs args)
        {
            if (args.Command == "А0")
                Format.A0();
        }
        protected override void OnCommand(Document document, int id)
        {
            bool flag;
            switch ((Commands)((id) % 11))
            {
                case Commands.myCommand1:
                    {
                        flag = Format.A4();
                        Replacing(id, flag);
                        break;
                    }
                case Commands.myCommand2:
                    {
                        flag = Format.A3();
                        Replacing(id, flag);
                        break;
                    }
                case Commands.myCommand3:
                    {
                        flag = Format.A2();
                        Replacing(id, flag);
                        break;
                    }
                case Commands.myCommand4:
                    {
                        flag = Format.A1();
                        Replacing(id, flag);
                        break;
                    }
                case Commands.myCommand5:
                    {
                        flag = Format.A0();
                        Replacing(id, flag);
                        break;
                    }
                case Commands.myCommand6:
                    {
                        flag = Format.A4x3();
                        Replacing(id, flag);
                        break;
                    }
                case Commands.myCommand7:
                    {
                        flag = Format.A3x3();
                        Replacing(id, flag);
                        break;
                    }
                case Commands.myCommand8:
                    {
                        flag = Format.A2x3();
                        Replacing(id, flag);
                        break;
                    }
                case Commands.myCommand9:
                    {
                        flag = Format.A1x3();
                        Replacing(id, flag);
                        break;
                    }
                case Commands.myCommand10:
                    {
                        flag = Format.A0x2();
                        Replacing(id, flag);
                        break;
                    }
                case Commands.myCommand11:
                    {
                        flag = Format.BCH();
                        Replacing(id, flag);
                        break;
                    }
                default:
                    base.OnCommand(document, id);
                    break;
            }
        }
        protected override void OnUpdateCommand(CommandUI cmdUI)
        {
            if (cmdUI == null)
                return;
            if (cmdUI.Document == null || cmdUI.Document.ActivePage == null)
            {
                cmdUI.Enable(false);
                return;
            }
            cmdUI.Enable();
        }
        protected override void NewDocumentCreatedEventHandler(DocumentEventArgs args)
        {
            args.Document.AttachPlugin(this);
        }
        protected override void DocumentOpenEventHandler(DocumentEventArgs args)
        {
            args.Document.AttachPlugin(this);
        }
    }
}