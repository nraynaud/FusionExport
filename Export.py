import os
import json

from adsk.core import *
import traceback

import tempfile
from shutil import copyfile

import sys

sys.path.append(os.path.join(os.path.dirname(__file__), 'libs'))

import Foundation

USER_PARAM = 'exporter_stl'
EXPORT_COMMAND_KEY = 'ExportDesign'
SETTINGS_COMMAND_KEY = 'SettingsExportDesign'

keepHandlers = []
deleters = []


def find_entity_from_path(path, component, ui):
    if len(path) > 1:
        if component.objectType == 'adsk::fusion::Component':
            child = component.occurrences.itemByName(path[0])
        else:
            child = component.childOccurrences.itemByName(path[0])
        if child:
            return find_entity_from_path(path[1:], child, ui)
    elif len(path) == 1:
        return component.bRepBodies.itemByName(path[0])
    return None


def get_path_for_entity(entity):
    chain = [entity.name]
    obj = entity
    while True:
        obj = obj.assemblyContext
        if obj:
            chain.append(obj.name)
        else:
            break
    return list(reversed(chain))


def select_output_file(ui):
    ui.createFileDialog()
    dialog = ui.createFileDialog()
    dialog.title = 'Select export file'
    dialog.filter = '*.stl'
    accessible = dialog.showSave()
    if accessible == DialogResults.DialogOK:
        return dialog.filename


def handle(handler, clazz):
    class Handler(clazz):
        def notify(self, args):
            try:
                handler(args)
            except:
                Application.get().userInterface.messageBox('Failed:\n{}'.format(traceback.format_exc()))

    handler_instance = Handler()
    keepHandlers.append(handler_instance)
    return handler_instance


def create_setting_panel(args):
    app = Application.get()
    ui = app.userInterface
    design = app.activeProduct
    cmd = Command.cast(args.command)

    def handle_file_input(args):
        event_args = InputChangedEventArgs.cast(args)
        cmd_input = event_args.input
        if cmd_input.id == 'selectFile':
            file_name = select_output_file(ui)
            if file_name:
                cmd_input.text = file_name

    cmd.inputChanged.add(handle(handle_file_input, InputChangedEventHandler))

    inputs = cmd.commandInputs
    cmd.okButtonText = 'Ok Export'
    selection_input = inputs.addSelectionInput('selection', 'Body', 'Basic select command input')
    selection_input.setSelectionLimits(1)
    selection_input.addSelectionFilter('Bodies')
    file_input = inputs.addBoolValueInput('selectFile', 'File', False, '', True)
    file_input.text = 'select file'

    def on_activation(args):
        dir_param = design.userParameters.itemByName(USER_PARAM)
        if dir_param:
            data = json.loads(dir_param.comment)
            body = find_entity_from_path(data['chain'], design.rootComponent, ui)
            if body:
                selection_input.addSelection(body)
            url, is_stale, error = Foundation.NSURL.URLByResolvingBookmarkData_options_relativeToURL_bookmarkDataIsStale_error_(
                Foundation.NSData.alloc().initWithBase64EncodedString_options_(data['file'], 0),
                Foundation.NSURLBookmarkResolutionWithSecurityScope,
                None,
                None,
                None)
            if not error:
                file_path = url.path()
                if file_path:
                    file_input.text = file_path

    cmd.activate.add((handle(on_activation, CommandEventHandler)))

    def on_execution(args):
        file_name = file_input.text
        body = selection_input.selection(0).entity
        chain = get_path_for_entity(body)
        url = Foundation.NSURL.alloc().initFileURLWithPath_(file_name)
        bookmark, error = url.bookmarkDataWithOptions_includingResourceValuesForKeys_relativeToURL_error_(
            Foundation.NSURLBookmarkCreationWithSecurityScope,
            None,
            None,
            None)
        # base64 encode the bookmark so it can be jsonified
        the_bytes = bookmark.base64EncodedStringWithOptions_(0)
        json_params = json.dumps({'file': the_bytes, 'chain': chain})
        dir_param = design.userParameters.itemByName(USER_PARAM)
        if dir_param:
            dir_param.comment = json_params
        else:
            design.userParameters.add(USER_PARAM, ValueInput.createByString('0'), '', json_params)
        ui.commandDefinitions.itemById(EXPORT_COMMAND_KEY).execute()

    cmd.execute.add(handle(on_execution, CommandEventHandler))


def export_stl_file(args):
    app = Application.get()
    ui = app.userInterface
    design = app.activeProduct
    dir_param = design.userParameters.itemByName(USER_PARAM)
    if dir_param:
        data = json.loads(dir_param.comment)
        body = find_entity_from_path(data['chain'], design.rootComponent, ui)
        if body:
            nsdata = Foundation.NSData.alloc().initWithBase64EncodedString_options_(data['file'], 0)
            url, is_stale, error = Foundation.NSURL.URLByResolvingBookmarkData_options_relativeToURL_bookmarkDataIsStale_error_(
                nsdata,
                Foundation.NSURLBookmarkResolutionWithSecurityScope,
                None,
                None,
                None)
            file_path = url.path()
            accessible = url.startAccessingSecurityScopedResource()
            if accessible:
                try:
                    # using the temp dir, for some reasons exportManager tries to access the directory
                    # surrounding the output file
                    with tempfile.NamedTemporaryFile(suffix='.stl') as temp_file:
                        export_manager = design.exportManager
                        options = export_manager.createSTLExportOptions(body, temp_file.name)
                        options.sendToPrintUtility = False
                        options.meshRefinement = 0
                        export_manager.execute(options)
                        temp_file.seek(0)
                        copyfile(temp_file.name, file_path)
                    return
                finally:
                    url.stopAccessingSecurityScopedResource()
            else:
                ui.messageBox('file ' + str(url) + ' is not accessible')
        else:
            ui.messageBox('body "' + '.'.join(data['chain']) + '" not found')
    ui.commandDefinitions.itemById(SETTINGS_COMMAND_KEY).execute()


def run(context):
    app = Application.get()
    ui = app.userInterface
    try:
        print('STARTING')
        cmd_defs = ui.commandDefinitions
        export_command = ui.commandDefinitions.itemById(EXPORT_COMMAND_KEY)
        if export_command:
            export_command.deleteMe()
        export_command = cmd_defs.addButtonDefinition(EXPORT_COMMAND_KEY, 'Export STL',
                                                      'Export the pre-determined STL file',
                                                      'exportIcon')
        settings_command = ui.commandDefinitions.itemById(SETTINGS_COMMAND_KEY)
        if settings_command:
            settings_command.deleteMe()
        settings_command = cmd_defs.addButtonDefinition(SETTINGS_COMMAND_KEY, 'Configure STL export',
                                                        'Configure STL export',
                                                        'configureExportIcon')
        export_command.commandCreated.add(handle(export_stl_file, CommandCreatedEventHandler))
        settings_command.commandCreated.add(handle(create_setting_panel, CommandCreatedEventHandler))
        deleters.append(
            replace_existing_control(ui.allToolbarPanels.itemById('SolidMakePanel').controls, settings_command))
        deleters.append(replace_existing_control(ui.toolbars.itemById('QAT').controls, export_command))
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def delete_control(controls, command_key):
    control = controls.itemById(command_key)
    if control:
        control.deleteMe()


def replace_existing_control(controls, new_command):
    delete_control(controls, new_command.id)
    if new_command:
        control = controls.addCommand(new_command)
        return lambda: control.deleteMe()
    return lambda: None


def stop(context):
    ui = None
    try:
        app = Application.get()
        ui = app.userInterface
        button = ui.commandDefinitions.itemById(EXPORT_COMMAND_KEY)
        if button:
            button.deleteMe()
        global keepHandlers
        keepHandlers = []
        for obj in deleters:
            obj()
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
