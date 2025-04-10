from flask_wtf import FlaskForm
from wtforms import SelectField, SubmitField

class DocumentPreferencesForm(FlaskForm):
    font_family = SelectField('Font Family', choices=[
        ('Arial', 'Arial'),
        ('Times New Roman', 'Times New Roman'),
        ('Helvetica', 'Helvetica'),
        ('Calibri', 'Calibri')
    ])
    font_size = SelectField('Font Size', choices=[
        ('10', '10pt'),
        ('11', '11pt'),
        ('12', '12pt'),
        ('14', '14pt')
    ])
    code_style = SelectField('Code Block Style', choices=[
        ('monokai', 'Monokai'),
        ('github', 'GitHub'),
        ('vs-code', 'VS Code'),
        ('dracula', 'Dracula')
    ])
    heading_style = SelectField('Heading Style', choices=[
        ('standard', 'Standard'),
        ('modern', 'Modern'),
        ('classic', 'Classic'),
        ('minimal', 'Minimal')
    ])
    submit = SubmitField('Save Preferences')