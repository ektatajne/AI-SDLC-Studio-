"""merge_heads

Revision ID: afeb7b03de11
Revises: d481f9a01e23, e582g1b02f34
Create Date: 2026-09-26 15:03:43.503346

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = 'afeb7b03de11'
down_revision: Union[str, Sequence[str], None] = ('d481f9a01e23', 'e582g1b02f34')
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    """Upgrade schema."""
    pass


def downgrade() -> None:
    """Downgrade schema."""
    pass
