#!/usr/bin/env python3
"""Build a proper .msapp file from source files."""
import json
import zipfile
import hashlib
import base64
import io
import os

# Template definitions for each control type used in the app
TEMPLATES = {
    "appinfo": {
        "Name": "appinfo",
        "Id": "http://microsoft.com/appmagic/appinfo",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
    "screen": {
        "Name": "screen",
        "Id": "http://microsoft.com/appmagic/screen",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
    "rectangle": {
        "Name": "rectangle",
        "Id": "http://microsoft.com/appmagic/rectangle",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
    "icon": {
        "Name": "icon",
        "Id": "http://microsoft.com/appmagic/icon",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
    "label": {
        "Name": "label",
        "Id": "http://microsoft.com/appmagic/label",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
    "image": {
        "Name": "image",
        "Id": "http://microsoft.com/appmagic/image",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
    "text": {
        "Name": "text",
        "Id": "http://microsoft.com/appmagic/text",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
    "button": {
        "Name": "button",
        "Id": "http://microsoft.com/appmagic/button",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
    "dropdown": {
        "Name": "dropdown",
        "Id": "http://microsoft.com/appmagic/dropdown",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
    "checkbox": {
        "Name": "checkbox",
        "Id": "http://microsoft.com/appmagic/checkbox",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
    "datePicker": {
        "Name": "datePicker",
        "Id": "http://microsoft.com/appmagic/datepicker",
        "Version": "1.0",
        "LastModifiedTimestamp": "0",
        "FirstParty": True,
        "IsComponent": False,
        "ComponentChangedSinceFileImport": False,
        "ComponentAllowCustomization": False,
        "ComponentExtraProperties": [],
        "IsComponentDefinition": False,
        "OverridableProperties": {}
    },
}


def make_rule(prop, script, provider="Unknown"):
    return {
        "Property": prop,
        "Category": "Data",
        "InvariantScript": script,
        "RuleProviderType": provider
    }


def make_control(name, uid, parent, template_name, rules, index=0, pub_order=0,
                 style="", children=None, variant=""):
    ctrl = {
        "Name": name,
        "ControlUniqueId": str(uid),
        "Parent": parent,
        "Index": index,
        "PublishOrderIndex": pub_order,
        "IsGroupControl": False,
        "VariantName": variant,
        "StyleName": style,
        "Template": TEMPLATES.get(template_name, {"Name": template_name}),
        "Rules": rules,
        "Children": children or []
    }
    return ctrl


def compute_checksum(data):
    """Compute Power Apps style checksum for a file."""
    h = hashlib.sha256(data).digest()
    return "C8_" + base64.b64encode(h).decode('ascii')


def build_msapp(output_path):
    # Header.json
    header = {
        "DocVersion": "1.326",
        "MinVersionToLoad": "1.326",
        "MSAppStructureVersion": "2.0",
        "LastSavedDateTimeUTC": "03/06/2026 12:00:00"
    }

    # Properties.json
    properties = {
        "Author": "",
        "Name": "CelebrationApp",
        "Id": "00000000-0000-0000-0000-000000000001",
        "FileID": "00000000-0000-0000-0000-000000000002",
        "DocumentAppType": "DesktopOrTablet",
        "DocumentLayoutWidth": 1366,
        "DocumentLayoutHeight": 768,
        "DocumentLayoutOrientation": "landscape",
        "DocumentLayoutScaleToFit": True,
        "DocumentLayoutMaintainAspectRatio": True,
        "DocumentLayoutLockOrientation": True,
        "DocumentType": "App",
        "AppCreationSource": "AppFromScratch",
        "AppDescription": "",
        "AppPreviewFlagsKey": [
            "delayloadscreens", "blockmovingcontrol", "projectionmapping",
            "usedisplaynamemetadata", "usenonblockingonstartrule",
            "useguiddatatypes", "useexperimentalcdsconnector",
            "useenforcesavedatalimits", "componentauthoring",
            "reliableconcurrent", "dataTableV2Control",
            "nativecdsexperimental", "useexperimentalsqlconnector",
            "enablecdsfileandlargeimage", "enhanceddelegation",
            "aibuilderserviceenrollment", "enablesummerlandgeospatialfeatures",
            "enablesummerlandmixedrealityfeatures"
        ],
        "DefaultConnectedDataSourceMaxGetRowsCount": 500,
        "InstrumentationKey": "",
        "LibraryDependencies": "[]",
        "LocalDatabaseReferences": "{}",
        "OriginatingVersion": "1.326",
        "ControlCount": {},
        "DeserializationLoadTime": 0,
        "AnalysisLoadTime": 0,
        "ErrorCount": 0
    }

    # PublishInfo.json
    publish_info = {
        "AppName": "CelebrationApp",
        "PublishTarget": "Player",
        "PublishResourcesLocally": False,
        "PublishDataLocally": False,
        "UserLocale": "en-US"
    }

    # Themes.json
    themes = {
        "CurrentTheme": "CelebrationTheme",
        "CustomThemes": [
            {
                "name": "CelebrationTheme",
                "palette": [
                    {"name": "themePrimary", "type": "c", "value": "RGBA(74, 25, 66, 1)"},
                    {"name": "themeLighterAlt", "type": "c", "value": "RGBA(247, 240, 246, 1)"},
                    {"name": "themeLighter", "type": "c", "value": "RGBA(224, 198, 221, 1)"},
                    {"name": "themeLight", "type": "c", "value": "RGBA(198, 153, 193, 1)"},
                    {"name": "themeTertiary", "type": "c", "value": "RGBA(146, 77, 138, 1)"},
                    {"name": "themeSecondary", "type": "c", "value": "RGBA(107, 31, 96, 1)"},
                    {"name": "themeDarkAlt", "type": "c", "value": "RGBA(67, 22, 59, 1)"},
                    {"name": "themeDark", "type": "c", "value": "RGBA(56, 18, 50, 1)"},
                    {"name": "themeDarker", "type": "c", "value": "RGBA(41, 13, 37, 1)"},
                    {"name": "neutralLighterAlt", "type": "c", "value": "RGBA(250, 249, 248, 1)"},
                    {"name": "neutralLighter", "type": "c", "value": "RGBA(243, 242, 241, 1)"},
                    {"name": "neutralLight", "type": "c", "value": "RGBA(237, 235, 233, 1)"},
                    {"name": "neutralQuaternaryAlt", "type": "c", "value": "RGBA(225, 223, 221, 1)"},
                    {"name": "neutralQuaternary", "type": "c", "value": "RGBA(208, 208, 208, 1)"},
                    {"name": "neutralTertiaryAlt", "type": "c", "value": "RGBA(200, 198, 196, 1)"},
                    {"name": "neutralTertiary", "type": "c", "value": "RGBA(161, 159, 157, 1)"},
                    {"name": "neutralSecondary", "type": "c", "value": "RGBA(96, 94, 92, 1)"},
                    {"name": "neutralPrimaryAlt", "type": "c", "value": "RGBA(59, 58, 57, 1)"},
                    {"name": "neutralPrimary", "type": "c", "value": "RGBA(50, 49, 48, 1)"},
                    {"name": "neutralDark", "type": "c", "value": "RGBA(32, 31, 30, 1)"},
                    {"name": "black", "type": "c", "value": "RGBA(0, 0, 0, 1)"},
                    {"name": "white", "type": "c", "value": "RGBA(255, 255, 255, 1)"}
                ],
                "styles": {}
            }
        ]
    }

    # DataSources.json
    data_sources = {"DataSources": []}

    # Templates.json with UsedTemplates
    templates_json = {
        "UsedTemplates": [
            {"Name": "appinfo", "Id": "http://microsoft.com/appmagic/appinfo", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
            {"Name": "screen", "Id": "http://microsoft.com/appmagic/screen", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
            {"Name": "rectangle", "Id": "http://microsoft.com/appmagic/rectangle", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
            {"Name": "icon", "Id": "http://microsoft.com/appmagic/icon", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
            {"Name": "label", "Id": "http://microsoft.com/appmagic/label", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
            {"Name": "image", "Id": "http://microsoft.com/appmagic/image", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
            {"Name": "text", "Id": "http://microsoft.com/appmagic/text", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
            {"Name": "button", "Id": "http://microsoft.com/appmagic/button", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
            {"Name": "dropdown", "Id": "http://microsoft.com/appmagic/dropdown", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
            {"Name": "checkbox", "Id": "http://microsoft.com/appmagic/checkbox", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
            {"Name": "datePicker", "Id": "http://microsoft.com/appmagic/datepicker", "Version": "1.0", "LastModifiedTimestamp": "0", "FirstParty": True},
        ]
    }

    # Resources.json
    resources = {"Resources": []}

    # --- Build Controls ---

    # OnStart script
    on_start = (
        'ClearCollect(\n'
        '    CelebrationEvents,\n'
        '    {\n'
        '        ID: 1,\n'
        '        CelebrationType: "Joiner",\n'
        '        EmployeeFirstName: "Michael",\n'
        '        EmployeeLastName: "Chen",\n'
        '        Status: "Needs Input",\n'
        '        ManagerName: "Vasilina Maroz",\n'
        '        TriggerDate: DateValue("2026-02-28"),\n'
        '        ManagerDraftEmail: "Dear Michal Chen, Welcome to the team! I am truly excited to have you join us and become a valued member of our group. My name is Vasilina MAROZ, and as your manager, I want to personally extend my warmest greetings. Our team is driven by core values of collaboration, innovation, and mutual support. We strive to create an environment where every member feels empowered to contribute their best and grow professionally. Your skills and enthusiasm will be a wonderful addition to our shared goals.",\n'
        '        CommunicationSender: "",\n'
        '        MyEmail: ""\n'
        '    }\n'
        ');\n'
        'ClearCollect(\n'
        '    SenderOptions,\n'
        '    {Value: "Select person who will send the communication..."},\n'
        '    {Value: "Vasilina Maroz"},\n'
        '    {Value: "HR Department"},\n'
        '    {Value: "Team Lead"}\n'
        ');\n'
        'Set(varCurrentEvent, LookUp(CelebrationEvents, ID = 1));\n'
        'Set(varShowMyEmail, true)'
    )

    # App control (Control 1)
    app_control = {
        "TopParent": make_control(
            name="App", uid=1, parent="", template_name="appinfo",
            style="defaultAppinfoStyle", pub_order=1,
            rules=[
                make_rule("BackEnabled", "true"),
                make_rule("OnStart", on_start),
                make_rule("StartScreen", "CelebrationDetailScreen"),
            ]
        )
    }

    # CelebrationDetailScreen children
    uid = 3  # screen uid
    children = []
    child_uid = 4

    def add_child(name, template, style, rules, variant=""):
        nonlocal child_uid
        c = make_control(
            name=name, uid=child_uid, parent="CelebrationDetailScreen",
            template_name=template, style=style,
            pub_order=child_uid, index=0,
            rules=rules, variant=variant
        )
        children.append(c)
        child_uid += 1

    add_child("rectHeader", "rectangle", "defaultRectangleStyle", [
        make_rule("X", "0"), make_rule("Y", "0"),
        make_rule("Width", "Parent.Width"), make_rule("Height", "80"),
        make_rule("Fill", "RGBA(74, 25, 66, 1)")
    ])

    add_child("icoBackArrow", "icon", "defaultIconStyle", [
        make_rule("X", "20"), make_rule("Y", "10"),
        make_rule("Width", "24"), make_rule("Height", "24"),
        make_rule("Color", "RGBA(255, 255, 255, 1)"),
        make_rule("OnSelect", "Back()"),
        make_rule("Icon", "Icon.BackArrow")
    ], variant="BackArrow")

    add_child("lblBackToActions", "label", "defaultLabelStyle", [
        make_rule("X", "48"), make_rule("Y", "8"),
        make_rule("Width", "200"), make_rule("Height", "28"),
        make_rule("Text", '"Back to My Actions"'),
        make_rule("Size", "12"), make_rule("Color", "RGBA(255, 255, 255, 1)"),
        make_rule("Font", "Font.'Open Sans'"), make_rule("OnSelect", "Back()")
    ])

    add_child("lblScreenTitle", "label", "defaultLabelStyle", [
        make_rule("X", "20"), make_rule("Y", "36"),
        make_rule("Width", "400"), make_rule("Height", "36"),
        make_rule("Text", '"Celebration Detail"'),
        make_rule("Size", "22"), make_rule("FontWeight", "FontWeight.Bold"),
        make_rule("Color", "RGBA(255, 255, 255, 1)"),
        make_rule("Font", "Font.'Open Sans'")
    ])

    add_child("imgUserAvatar", "image", "defaultImageStyle", [
        make_rule("X", "Parent.Width - 60"), make_rule("Y", "20"),
        make_rule("Width", "40"), make_rule("Height", "40"),
        make_rule("BorderRadius", "20"),
        make_rule("ImagePosition", "ImagePosition.Fill"),
        make_rule("Image", "SampleImage")
    ])

    add_child("lblUserName", "label", "defaultLabelStyle", [
        make_rule("X", "Parent.Width - 220"), make_rule("Y", "28"),
        make_rule("Width", "150"), make_rule("Height", "28"),
        make_rule("Text", '"Vasilina Maroz"'),
        make_rule("Size", "13"), make_rule("Color", "RGBA(255, 255, 255, 1)"),
        make_rule("Align", "Align.Right"), make_rule("Font", "Font.'Open Sans'")
    ])

    add_child("rectInfoBanner", "rectangle", "defaultRectangleStyle", [
        make_rule("X", "40"), make_rule("Y", "95"),
        make_rule("Width", "Parent.Width - 80"), make_rule("Height", "55"),
        make_rule("Fill", "RGBA(255, 243, 224, 1)"),
        make_rule("BorderColor", "RGBA(245, 124, 0, 1)"),
        make_rule("BorderThickness", "1"), make_rule("BorderRadius", "4")
    ])

    add_child("lblInfoText", "label", "defaultLabelStyle", [
        make_rule("X", "55"), make_rule("Y", "100"),
        make_rule("Width", "Parent.Width - 110"), make_rule("Height", "45"),
        make_rule("Text", '"The email text is pre-created for you using a template prompt. You can edit the text directly in the textbox if needed. You can also click Copy template prompt to copy the prompt and use Tell Me to generate another message."'),
        make_rule("Size", "11"), make_rule("Color", "RGBA(230, 81, 0, 1)"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("VerticalAlign", "VerticalAlign.Middle"),
        make_rule("Overflow", "Overflow.Hidden")
    ])

    add_child("lblEventSummaryHeader", "label", "defaultLabelStyle", [
        make_rule("X", "55"), make_rule("Y", "170"),
        make_rule("Width", "220"), make_rule("Height", "35"),
        make_rule("Text", '"Event Summary"'),
        make_rule("Size", "18"), make_rule("FontWeight", "FontWeight.Bold"),
        make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'")
    ])

    add_child("lblCelebrationTypeLabel", "label", "defaultLabelStyle", [
        make_rule("X", "55"), make_rule("Y", "215"),
        make_rule("Width", "180"), make_rule("Height", "22"),
        make_rule("Text", '"Celebration Type"'),
        make_rule("Size", "11"), make_rule("Color", "RGBA(102, 102, 102, 1)"),
        make_rule("Font", "Font.'Open Sans'")
    ])

    add_child("btnCelebrationTypeBadge", "button", "defaultButtonStyle", [
        make_rule("X", "55"), make_rule("Y", "240"),
        make_rule("Width", "80"), make_rule("Height", "30"),
        make_rule("Text", '"Joiner"'),
        make_rule("Size", "11"), make_rule("Color", "RGBA(255, 255, 255, 1)"),
        make_rule("Fill", "RGBA(0, 137, 123, 1)"),
        make_rule("BorderColor", "RGBA(0, 137, 123, 1)"),
        make_rule("BorderRadius", "15"), make_rule("Font", "Font.'Open Sans'"),
        make_rule("FontWeight", "FontWeight.Semibold"),
        make_rule("DisabledFill", "RGBA(0, 137, 123, 1)"),
        make_rule("DisabledColor", "RGBA(255, 255, 255, 1)"),
        make_rule("DisplayMode", "DisplayMode.Disabled")
    ])

    add_child("lblFirstNameLabel", "label", "defaultLabelStyle", [
        make_rule("X", "55"), make_rule("Y", "285"),
        make_rule("Width", "180"), make_rule("Height", "22"),
        make_rule("Text", '"Employee First Name"'),
        make_rule("Size", "11"), make_rule("Color", "RGBA(102, 102, 102, 1)"),
        make_rule("Font", "Font.'Open Sans'")
    ])

    add_child("lblFirstNameValue", "label", "defaultLabelStyle", [
        make_rule("X", "55"), make_rule("Y", "307"),
        make_rule("Width", "200"), make_rule("Height", "28"),
        make_rule("Text", "varCurrentEvent.EmployeeFirstName"),
        make_rule("Size", "15"), make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'")
    ])

    add_child("lblLastNameLabel", "label", "defaultLabelStyle", [
        make_rule("X", "55"), make_rule("Y", "350"),
        make_rule("Width", "180"), make_rule("Height", "22"),
        make_rule("Text", '"Employee Last Name"'),
        make_rule("Size", "11"), make_rule("Color", "RGBA(102, 102, 102, 1)"),
        make_rule("Font", "Font.'Open Sans'")
    ])

    add_child("lblLastNameValue", "label", "defaultLabelStyle", [
        make_rule("X", "55"), make_rule("Y", "372"),
        make_rule("Width", "200"), make_rule("Height", "28"),
        make_rule("Text", "varCurrentEvent.EmployeeLastName"),
        make_rule("Size", "15"), make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'")
    ])

    add_child("lblStatusLabel", "label", "defaultLabelStyle", [
        make_rule("X", "55"), make_rule("Y", "415"),
        make_rule("Width", "180"), make_rule("Height", "22"),
        make_rule("Text", '"Current Status"'),
        make_rule("Size", "11"), make_rule("Color", "RGBA(102, 102, 102, 1)"),
        make_rule("Font", "Font.'Open Sans'")
    ])

    add_child("lblStatusValue", "label", "defaultLabelStyle", [
        make_rule("X", "55"), make_rule("Y", "437"),
        make_rule("Width", "200"), make_rule("Height", "28"),
        make_rule("Text", "varCurrentEvent.Status"),
        make_rule("Size", "13"), make_rule("Color", "RGBA(216, 67, 21, 1)"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("FontWeight", "FontWeight.Semibold"),
        make_rule("Underline", "true")
    ])

    add_child("rectDivider", "rectangle", "defaultRectangleStyle", [
        make_rule("X", "300"), make_rule("Y", "170"),
        make_rule("Width", "1"), make_rule("Height", "480"),
        make_rule("Fill", "RGBA(224, 224, 224, 1)")
    ])

    add_child("lblRequiredInputHeader", "label", "defaultLabelStyle", [
        make_rule("X", "330"), make_rule("Y", "170"),
        make_rule("Width", "220"), make_rule("Height", "35"),
        make_rule("Text", '"Required Input"'),
        make_rule("Size", "18"), make_rule("FontWeight", "FontWeight.Bold"),
        make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'")
    ])

    add_child("lblManagerEmailLabel", "label", "defaultLabelStyle", [
        make_rule("X", "330"), make_rule("Y", "215"),
        make_rule("Width", "200"), make_rule("Height", "22"),
        make_rule("Text", '"Manager Draft Email"'),
        make_rule("Size", "12"), make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("FontWeight", "FontWeight.Semibold")
    ])

    add_child("lblCopyPrompt", "label", "defaultLabelStyle", [
        make_rule("X", "Parent.Width - 190"), make_rule("Y", "215"),
        make_rule("Width", "150"), make_rule("Height", "22"),
        make_rule("Text", '"Copy template prompt"'),
        make_rule("Size", "11"), make_rule("Color", "RGBA(74, 25, 66, 1)"),
        make_rule("Font", "Font.'Open Sans'"), make_rule("Align", "Align.Right"),
        make_rule("Underline", "true"),
        make_rule("OnSelect",
            'Copy(\n'
            '    "Write a warm welcome email from " & varCurrentEvent.ManagerName &\n'
            '    " to " & varCurrentEvent.EmployeeFirstName & " " & varCurrentEvent.EmployeeLastName &\n'
            '    " who is joining the team."\n'
            ')')
    ])

    add_child("txtManagerEmail", "text", "defaultTextStyle", [
        make_rule("X", "330"), make_rule("Y", "242"),
        make_rule("Width", "Parent.Width - 370"), make_rule("Height", "160"),
        make_rule("Default", "varCurrentEvent.ManagerDraftEmail"),
        make_rule("Mode", "TextMode.MultiLine"),
        make_rule("Size", "12"), make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("BorderColor", "RGBA(189, 189, 189, 1)"),
        make_rule("BorderThickness", "1"), make_rule("BorderRadius", "4"),
        make_rule("Fill", "RGBA(255, 255, 255, 1)"),
        make_rule("PaddingTop", "10"), make_rule("PaddingBottom", "10"),
        make_rule("PaddingLeft", "12"), make_rule("PaddingRight", "12")
    ])

    add_child("lblTriggerDateLabel", "label", "defaultLabelStyle", [
        make_rule("X", "330"), make_rule("Y", "415"),
        make_rule("Width", "250"), make_rule("Height", "22"),
        make_rule("Text", '"Communication Trigger Date"'),
        make_rule("Size", "12"), make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("FontWeight", "FontWeight.Semibold")
    ])

    add_child("dpTriggerDate", "datePicker", "defaultDatePickerStyle", [
        make_rule("X", "330"), make_rule("Y", "440"),
        make_rule("Width", "220"), make_rule("Height", "38"),
        make_rule("DefaultDate", "varCurrentEvent.TriggerDate"),
        make_rule("Size", "12"), make_rule("Font", "Font.'Open Sans'"),
        make_rule("BorderColor", "RGBA(189, 189, 189, 1)"),
        make_rule("BorderThickness", "1"), make_rule("BorderRadius", "4"),
        make_rule("IconFill", "RGBA(102, 102, 102, 1)")
    ])

    add_child("lblSenderLabel", "label", "defaultLabelStyle", [
        make_rule("X", "330"), make_rule("Y", "490"),
        make_rule("Width", "250"), make_rule("Height", "22"),
        make_rule("Text", '"Communication Sender"'),
        make_rule("Size", "12"), make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("FontWeight", "FontWeight.Semibold")
    ])

    add_child("ddSender", "dropdown", "defaultDropdownStyle", [
        make_rule("X", "330"), make_rule("Y", "515"),
        make_rule("Width", "Parent.Width - 370"), make_rule("Height", "38"),
        make_rule("Items", "SenderOptions"),
        make_rule("Size", "12"), make_rule("Font", "Font.'Open Sans'"),
        make_rule("BorderColor", "RGBA(189, 189, 189, 1)"),
        make_rule("BorderThickness", "1"), make_rule("BorderRadius", "4"),
        make_rule("ChevronFill", "RGBA(102, 102, 102, 1)"),
        make_rule("Fill", "RGBA(255, 255, 255, 1)")
    ])

    add_child("chkSendOwnEmail", "checkbox", "defaultCheckboxStyle", [
        make_rule("X", "330"), make_rule("Y", "567"),
        make_rule("Width", "350"), make_rule("Height", "30"),
        make_rule("Text", '"I want to send my own email as well"'),
        make_rule("Size", "12"), make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("CheckboxBackgroundFill", "RGBA(255, 255, 255, 1)"),
        make_rule("CheckboxBorderColor", "RGBA(74, 25, 66, 1)"),
        make_rule("CheckmarkFill", "RGBA(74, 25, 66, 1)"),
        make_rule("Default", "true"),
        make_rule("OnCheck", "Set(varShowMyEmail, true)"),
        make_rule("OnUncheck", "Set(varShowMyEmail, false)")
    ])

    add_child("lblMyEmailLabel", "label", "defaultLabelStyle", [
        make_rule("X", "330"), make_rule("Y", "605"),
        make_rule("Width", "120"), make_rule("Height", "22"),
        make_rule("Text", '"My Email"'),
        make_rule("Size", "12"), make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("FontWeight", "FontWeight.Semibold"),
        make_rule("Visible", "varShowMyEmail")
    ])

    add_child("btnDraftMyEmail", "button", "defaultButtonStyle", [
        make_rule("X", "Parent.Width - 190"), make_rule("Y", "600"),
        make_rule("Width", "150"), make_rule("Height", "32"),
        make_rule("Text", '"✏ Draft my Email"'),
        make_rule("Size", "11"), make_rule("Color", "RGBA(74, 25, 66, 1)"),
        make_rule("Fill", "RGBA(255, 255, 255, 1)"),
        make_rule("BorderColor", "RGBA(74, 25, 66, 1)"),
        make_rule("BorderThickness", "1"), make_rule("BorderRadius", "4"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("Visible", "varShowMyEmail"),
        make_rule("OnSelect", 'Notify("Drafting your email...", NotificationType.Information)')
    ])

    add_child("txtMyEmail", "text", "defaultTextStyle", [
        make_rule("X", "330"), make_rule("Y", "635"),
        make_rule("Width", "Parent.Width - 370"), make_rule("Height", "80"),
        make_rule("Default", '""'),
        make_rule("HintText", '"Enter your personal email message..."'),
        make_rule("Mode", "TextMode.MultiLine"),
        make_rule("Size", "12"), make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("BorderColor", "RGBA(189, 189, 189, 1)"),
        make_rule("BorderThickness", "1"), make_rule("BorderRadius", "4"),
        make_rule("Fill", "RGBA(255, 255, 255, 1)"),
        make_rule("PaddingTop", "10"), make_rule("PaddingLeft", "12"),
        make_rule("Visible", "varShowMyEmail")
    ])

    add_child("btnDeclineEvent", "button", "defaultButtonStyle", [
        make_rule("X", "Parent.Width - 500"), make_rule("Y", "Parent.Height - 55"),
        make_rule("Width", "140"), make_rule("Height", "40"),
        make_rule("Text", '"Decline Event"'),
        make_rule("Size", "13"), make_rule("Color", "RGBA(211, 47, 47, 1)"),
        make_rule("Fill", "RGBA(255, 255, 255, 1)"),
        make_rule("BorderColor", "RGBA(211, 47, 47, 1)"),
        make_rule("BorderThickness", "2"), make_rule("BorderRadius", "4"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("FontWeight", "FontWeight.Semibold"),
        make_rule("HoverFill", "RGBA(255, 235, 238, 1)"),
        make_rule("PressedFill", "RGBA(255, 205, 210, 1)"),
        make_rule("OnSelect", 'Notify("Event declined", NotificationType.Warning);\nBack()')
    ])

    add_child("btnSaveDraft", "button", "defaultButtonStyle", [
        make_rule("X", "Parent.Width - 340"), make_rule("Y", "Parent.Height - 55"),
        make_rule("Width", "130"), make_rule("Height", "40"),
        make_rule("Text", '"Save Draft"'),
        make_rule("Size", "13"), make_rule("Color", "RGBA(51, 51, 51, 1)"),
        make_rule("Fill", "RGBA(255, 255, 255, 1)"),
        make_rule("BorderColor", "RGBA(158, 158, 158, 1)"),
        make_rule("BorderThickness", "2"), make_rule("BorderRadius", "4"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("FontWeight", "FontWeight.Semibold"),
        make_rule("HoverFill", "RGBA(245, 245, 245, 1)"),
        make_rule("PressedFill", "RGBA(224, 224, 224, 1)"),
        make_rule("OnSelect",
            'Patch(\n'
            '    CelebrationEvents,\n'
            '    varCurrentEvent,\n'
            '    {\n'
            '        ManagerDraftEmail: txtManagerEmail.Text,\n'
            '        TriggerDate: dpTriggerDate.SelectedDate,\n'
            '        MyEmail: txtMyEmail.Text\n'
            '    }\n'
            ');\n'
            'Notify("Draft saved successfully", NotificationType.Success)')
    ])

    add_child("btnSubmitForReview", "button", "defaultButtonStyle", [
        make_rule("X", "Parent.Width - 190"), make_rule("Y", "Parent.Height - 55"),
        make_rule("Width", "160"), make_rule("Height", "40"),
        make_rule("Text", '"Submit For Review"'),
        make_rule("Size", "13"), make_rule("Color", "RGBA(255, 255, 255, 1)"),
        make_rule("Fill", "RGBA(45, 45, 45, 1)"),
        make_rule("BorderColor", "RGBA(45, 45, 45, 1)"),
        make_rule("BorderThickness", "0"), make_rule("BorderRadius", "4"),
        make_rule("Font", "Font.'Open Sans'"),
        make_rule("FontWeight", "FontWeight.Semibold"),
        make_rule("HoverFill", "RGBA(66, 66, 66, 1)"),
        make_rule("PressedFill", "RGBA(33, 33, 33, 1)"),
        make_rule("OnSelect",
            'Patch(\n'
            '    CelebrationEvents,\n'
            '    varCurrentEvent,\n'
            '    {\n'
            '        Status: "Submitted",\n'
            '        ManagerDraftEmail: txtManagerEmail.Text,\n'
            '        TriggerDate: dpTriggerDate.SelectedDate,\n'
            '        MyEmail: txtMyEmail.Text\n'
            '    }\n'
            ');\n'
            'Notify("Submitted for review", NotificationType.Success);\n'
            'Back()')
    ])

    # Screen control (Control 3)
    screen_control = {
        "TopParent": make_control(
            name="CelebrationDetailScreen", uid=3, parent="",
            template_name="screen", style="defaultScreenStyle", pub_order=2,
            rules=[make_rule("Fill", "RGBA(255, 255, 255, 1)")],
            children=children
        )
    }

    # --- Assemble ZIP ---
    files = {}
    files["Header.json"] = json.dumps(header, indent=2).encode('utf-8')
    files["Properties.json"] = json.dumps(properties, indent=2).encode('utf-8')
    files["Resources\\PublishInfo.json"] = json.dumps(publish_info, indent=2).encode('utf-8')
    files["References\\Themes.json"] = json.dumps(themes, indent=2).encode('utf-8')
    files["References\\DataSources.json"] = json.dumps(data_sources, indent=2).encode('utf-8')
    files["References\\Templates.json"] = json.dumps(templates_json, indent=2).encode('utf-8')
    files["References\\Resources.json"] = json.dumps(resources, indent=2).encode('utf-8')
    files["Controls\\1.json"] = json.dumps(app_control, indent=2).encode('utf-8')
    files["Controls\\3.json"] = json.dumps(screen_control, indent=2).encode('utf-8')

    # Compute checksums
    per_file = {}
    for name, data in files.items():
        per_file[name] = compute_checksum(data)

    # Compute overall checksum from sorted per-file checksums
    combined = "".join(per_file[k] for k in sorted(per_file.keys()))
    overall = compute_checksum(combined.encode('utf-8'))

    checksum = {
        "ClientStampedChecksum": overall,
        "ClientPerFileChecksums": per_file,
        "ServerStampedChecksum": ""
    }
    files["checksum.json"] = json.dumps(checksum, indent=2).encode('utf-8')

    # Write ZIP
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)

    print(f"Created {output_path}")
    print(f"Files: {len(files)}")
    for name in files:
        print(f"  {name} ({len(files[name])} bytes)")


if __name__ == "__main__":
    build_msapp("/home/user/Celebration-app/CelebrationApp.msapp")
