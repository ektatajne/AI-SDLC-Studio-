"""add_testing_version_table

Revision ID: e582g1b02f34
Revises: d481f9a01e23
Create Date: 2026-09-26 14:18:00.000000

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = 'e582g1b02f34'
down_revision: Union[str, Sequence[str], None] = '218783eda80c'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table('testing_versions',
        sa.Column('id', sa.String(length=36), nullable=False),
        sa.Column('project_id', sa.String(length=36), nullable=False),
        sa.Column('development_version_id', sa.String(length=36), nullable=True),
        sa.Column('version_num', sa.Integer(), nullable=True),
        sa.Column('raw_execution_response', sa.Text(), nullable=True),
        sa.Column('raw_report_response', sa.Text(), nullable=True),
        sa.Column('approval_status', sa.String(length=50), nullable=True),
        sa.Column('quality_gate_status', sa.String(length=50), nullable=True),
        sa.Column('reviewer_comments', sa.Text(), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
        sa.ForeignKeyConstraint(['project_id'], ['projects.id'], ),
        sa.ForeignKeyConstraint(['development_version_id'], ['development_versions.id'], ),
        sa.PrimaryKeyConstraint('id')
    )


def downgrade() -> None:
    op.drop_table('testing_versions')
