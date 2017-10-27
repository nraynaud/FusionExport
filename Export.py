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
    dialog = ui.createFileDialog()
    dialog.title = 'Export STL file'
    dialog.filter = 'STL files (*.stl)'
    dialog.initialFilename = 'out'
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


def get_export_list():
    attributes = Application.get().activeProduct.findAttributes('nraynaud-Export', 'export')
    result = []
    for attr in attributes:
        body = attr.parent
        content = json.loads(attr.value)
        if body:
            for entry in content:
                result.append({'body': body, 'file': entry['file']})
    return result


def create_setting_panel(args):
    app = Application.get()
    ui = app.userInterface
    design = app.activeProduct
    cmd = Command.cast(args.command)

    shadow_table = []
    deletion_table = []
    shadow_dict = {}
    current_file = None

    def select_row(new_row):
        nonlocal current_file
        set_detail_visibility(True)
        body = new_row.get('body')
        if body:
            selection_input.addSelection(body)
        file_input.text = str(new_row.get('file'))
        current_file = new_row.get('file')

    def save_if_you_can():
        if selection_input.selectionCount and current_file is not None:
            if table.selectedRow == -1:
                row_index = table.rowCount
                row_id = 'TableInput_string{}'.format(row_index)
                body = selection_input.selection(0).entity
                shadow_row = {'id': row_id, 'file': current_file, 'body': body, 'new': True}
                shadow_table.append(shadow_row)
                shadow_dict[row_id] = shadow_row
                string_input = table.commandInputs.addStringValueInput(row_id, 'String', current_file)
                string_input.isReadOnly = True
                shadow_row['input'] = string_input
                table.addCommandInput(string_input, row_index, 0)
                table.selectedRow = row_index
            else:
                shadow_row = shadow_table[table.selectedRow]
                # adding the old row to deletion list will remove the attribute from the element
                deletion_table.append(shadow_row.copy())
                shadow_row['file'] = current_file
                shadow_row['body'] = selection_input.selection(0).entity

    def set_detail_visibility(visible):
        selection_input.clearSelection()
        if visible:
            selection_input.isVisible = True
            file_input.isVisible = True
        else:
            selection_input.isVisible = False
            file_input.isVisible = False

    def handle_input_change(args):
        nonlocal current_file
        event_args = InputChangedEventArgs.cast(args)
        cmd_input = event_args.input
        if cmd_input.id == 'selection':
            save_if_you_can()
        if cmd_input.id == 'selectFile':
            file_name = select_output_file(ui)
            if file_name:
                current_file = file_name
                cmd_input.text = file_name
                save_if_you_can()
        if cmd_input.id == 'addExport':
            table.selectedRow = -1
            remove_export.isEnabled = False
            current_file = None
            file_input.text = 'select file'
            set_detail_visibility(True)
        if cmd_input.id == 'removeExport':
            shadow_row = shadow_table[table.selectedRow]
            deletion_table.append(shadow_row)
            del shadow_table[table.selectedRow]
            del shadow_dict[shadow_row['id']]
            table.deleteRow(table.selectedRow)
            table.selectedRow = -1
            remove_export.isEnabled = False
            current_file = None
            file_input.text = 'select file'
            set_detail_visibility(False)
        else:
            shadow_row = shadow_dict.get(cmd_input.id)
            if shadow_row:
                select_row(shadow_row)
                remove_export.isEnabled = True

    inputs = cmd.commandInputs
    table = inputs.addTableCommandInput('sampleTable', 'Table', 2, '1:1')
    add_export = inputs.addBoolValueInput('addExport', 'Add Export', False, '', True)
    add_export.text = '+'
    add_export.tooltip = 'Add an export'
    table.addToolbarCommandInput(add_export)
    remove_export = inputs.addBoolValueInput('removeExport', 'Remove Export', False, '', True)
    remove_export.text = '-'
    remove_export.tooltip = "Don't export this anymore"
    remove_export.isEnabled = False
    table.addToolbarCommandInput(remove_export)
    cmd.okButtonText = 'Ok Export'
    selection_input = inputs.addSelectionInput('selection', 'Body', 'Basic select command input')
    selection_input.addSelectionFilter('Bodies')
    # allow 0 selection so that it doesn't invalidate the entire command if empty
    selection_input.setSelectionLimits(0, 1)
    file_input = inputs.addBoolValueInput('selectFile', 'File', False, '', True)
    file_input.text = 'select file'
    set_detail_visibility(False)

    def on_activation(args):
        exports = get_export_list()
        index = 0
        for export in exports:
            file_path = decode_bookmark(export['file'])
            row_id = 'TableInput_string{}'.format(index)
            shadow_row = {'id': row_id, 'file': file_path, 'body': export['body']}
            shadow_table.append(shadow_row)
            shadow_dict[shadow_row['id']] = shadow_row
            string_input = table.commandInputs.addStringValueInput(row_id, 'String', file_path)
            shadow_row['input'] = string_input
            string_input.isReadOnly = True
            table.addCommandInput(string_input, index, 0)
            index += 1

    def add_or_replace_attribute(element, new_value):
        attr = element.attributes.itemByName('nraynaud-Export', 'export')
        if attr:
            previous_content = list(json.loads(attr.value))
            if not next((e for e in previous_content if e['file'] == new_value['file']), None):
                previous_content.append(new_value)
                attr.value = json.dumps(previous_content)
        else:
            element.attributes.add('nraynaud-Export', 'export', str(json.dumps([new_value])))

    def on_execution(args):
        for row in deletion_table:
            file_name = row['file']
            body = row['body']
            the_bytes = get_bookmark_bytes(file_name)
            attr = body.attributes.itemByName('nraynaud-Export', 'export')
            if attr:
                previous_content = json.loads(attr.value)
                left_content = [e for e in previous_content if e['file'] != the_bytes]
                if len(left_content):
                    attr.value = json.dumps(left_content)
                else:
                    attr.deleteMe()
        for row in shadow_table:
            file_name = row['file']
            body = row['body']
            the_bytes = get_bookmark_bytes(file_name)
            need_save = True
            if not the_bytes:
                export_stl(design, body, file_name)
                the_bytes = get_bookmark_bytes(file_name)
                need_save = False
            add_or_replace_attribute(body, {'file': the_bytes, 'type': 'stl'})
            if need_save:
                export_to_bookmark(the_bytes, design, body)

    cmd.inputChanged.add(handle(handle_input_change, InputChangedEventHandler))
    cmd.activate.add((handle(on_activation, CommandEventHandler)))
    cmd.execute.add(handle(on_execution, CommandEventHandler))


def get_bookmark_bytes(file_name):
    url = Foundation.NSURL.alloc().initFileURLWithPath_(file_name)
    bookmark, error = url.bookmarkDataWithOptions_includingResourceValuesForKeys_relativeToURL_error_(
        Foundation.NSURLBookmarkCreationWithSecurityScope,
        None,
        None,
        None)
    if error:
        # that's the error code if the file is not on disk yet.
        if error.code() == 260:
            return None
        raise Exception(error.localizedDescription())
    # base64 encode the bookmark so it can be jsonified
    return bookmark.base64EncodedStringWithOptions_(0)


def decode_bookmark(bookmark, access_action=None):
    nsdata = Foundation.NSData.alloc().initWithBase64EncodedString_options_(bookmark, 0)
    url, is_stale, error = Foundation.NSURL.URLByResolvingBookmarkData_options_relativeToURL_bookmarkDataIsStale_error_(
        nsdata, Foundation.NSURLBookmarkResolutionWithSecurityScope, None, None, None)
    if error:
        raise Exception(error.localizedDescription())
    file_path = url.path()
    if access_action:
        accessible = url.startAccessingSecurityScopedResource()
        if accessible:
            try:
                return access_action(file_path)
            finally:
                url.stopAccessingSecurityScopedResource()
    else:
        return file_path


def export_to_bookmark(bookmark, design, body):
    return decode_bookmark(bookmark, lambda file_path: export_stl(design, body, file_path))


def export_all_files(args):
    exports = get_export_list()
    design = Application.get().activeProduct
    for export in exports:
        export_to_bookmark(export['file'], design, export['body'])


def export_stl(design, body, file_path):
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


def run(context):
    app = Application.get()
    ui = app.userInterface
    try:
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
        export_command.commandCreated.add(handle(export_all_files, CommandCreatedEventHandler))
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
