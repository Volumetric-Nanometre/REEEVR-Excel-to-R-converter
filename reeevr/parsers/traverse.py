import reeevr.parsers.excelast as excelast
from reeevr.converters.variable import VariableConverter


class TraverseTree:
    def __init__(self, excelast, sheet, coordinate):

        self.tree = excelast.body
        self.depth = 0
        self.outputsheet = sheet
        self.outputcoordinate = coordinate
        self.outputvarname = VariableConverter.excel_cell_to_variable(sheet, coordinate)
        self.count = 0
        self.code = ""
        self.function_transformations = {}
        self.variables = []

    def walk(self, tree):

        simplesyntax = []
        for node in tree:
            self._ast_transform(node)
            if isinstance(node, excelast.ExpressionNode):
                self.depth += 1
                self.count += 1

                simplesyntax.append(self._transform(node))

                self.depth -= 1
            elif node.type == 'Separator':
                simplesyntax.append('%sep%')
            else:
                simplesyntax.append(node.value)

        if self.depth == 0:
            self.code += f"{self.outputvarname} = "+" ".join(simplesyntax)
        return simplesyntax

    def _transform(self, node):
        """
        Take node and select correct function transformation and
        builder
        :param node: ExpressionNode node from AST
        :return: code/variable name representing the formula
        """

        try:
            return self.function_transformations[node.name](node.params)
        except KeyError:
            raise KeyError("Function not in transform list")

    def _ast_transform(self, node):
        """
        Replace node values with variable names
        and lists
        :param node:
        :return:
        """
        if isinstance(node, excelast.OperandNode) and node.subtype == 'RANGE':

            sheet = self.outputsheet
            coordinate = node.value
            if '!' in node.value:
                sheet, coordinate = node.value.split('!')

            if ':' in node.value:
                code,  variables = VariableConverter.excel_range_to_list(sheet, coordinate)
                self.variables += variables
                node.value = code
            else:
                variable = VariableConverter.excel_cell_to_variable(sheet, coordinate)
                self.variables.append(variable)
                node.value = variable
        else:
            pass
