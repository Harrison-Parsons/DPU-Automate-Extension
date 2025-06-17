// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using Microsoft.CommandPalette.Extensions;
using Microsoft.CommandPalette.Extensions.Toolkit;
using UtilitiesAutomateExtension.Pages;

namespace UtilitiesAutomateExtension;

public partial class UtilitiesAutomateExtensionCommandsProvider : CommandProvider
{
    private readonly ICommandItem[] _commands;

    public UtilitiesAutomateExtensionCommandsProvider()
    {
        ContactSearch contactSearch = new ContactSearch();

        DisplayName = "DPU Automation";
        Icon = new IconInfo("\uE99A");
        _commands = [
            new CommandItem(new UtilitiesAutomateExtensionPage()) { Title = DisplayName, Icon = new IconInfo("\uE99A"),},
            new CommandItem(new ContactSearch()) { Title = "Contact Search", Icon = new IconInfo("\ue716") },
        ];
    }

    public override ICommandItem[] TopLevelCommands()
    {
        return _commands;
    }

}
