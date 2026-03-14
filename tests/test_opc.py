"""Tests for OPC package layer."""

from __future__ import annotations

import io

from slidecraft.opc.content_types import ContentTypeMap
from slidecraft.opc.package import OpcPackage
from slidecraft.opc.part import Part
from slidecraft.opc.relationships import Relationship, RelationshipCollection
from slidecraft.util.color import RGBColor
from slidecraft.util.units import Cm, Emu, Inches, Pt
from slidecraft.xml.ns import qn


class TestContentTypeMap:
    def test_parse_and_lookup(self) -> None:
        xml = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            b'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            b'<Default Extension="xml" ContentType="application/xml"/>'
            b'<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
            b"</Types>"
        )
        ct = ContentTypeMap.from_xml(xml)
        assert ct.content_type_for("/ppt/presentation.xml") is not None
        assert "presentation" in (ct.content_type_for("/ppt/presentation.xml") or "")
        assert ct.content_type_for("/some/file.rels") is not None
        assert ct.content_type_for("/some/file.xml") == "application/xml"

    def test_add_and_serialize(self) -> None:
        ct = ContentTypeMap()
        ct.add_default("xml", "application/xml")
        ct.add_override("/ppt/slides/slide1.xml", "application/slide")
        xml_bytes = ct.to_xml()
        ct2 = ContentTypeMap.from_xml(xml_bytes)
        assert ct2.content_type_for("/ppt/slides/slide1.xml") == "application/slide"
        assert ct2.content_type_for("/any.xml") == "application/xml"


class TestRelationshipCollection:
    def test_parse_and_lookup(self) -> None:
        xml = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" Type="http://type/a" Target="target.xml"/>'
            b'<Relationship Id="rId2" Type="http://type/b" Target="http://ext" TargetMode="External"/>'
            b"</Relationships>"
        )
        rels = RelationshipCollection.from_xml(xml)
        assert len(rels) == 2
        r1 = rels.get_by_id("rId1")
        assert r1 is not None
        assert r1.target == "target.xml"
        assert not r1.is_external
        r2 = rels.get_by_id("rId2")
        assert r2 is not None
        assert r2.is_external

    def test_add_and_serialize(self) -> None:
        rels = RelationshipCollection.new()
        r = rels.add("http://type/slide", "slides/slide1.xml")
        assert r.r_id == "rId1"
        xml_bytes = rels.to_xml()
        rels2 = RelationshipCollection.from_xml(xml_bytes)
        assert len(rels2) == 1

    def test_get_by_type(self) -> None:
        rels = RelationshipCollection.new()
        rels.add("http://type/a", "a.xml")
        rels.add("http://type/b", "b.xml")
        rels.add("http://type/a", "a2.xml")
        found = rels.get_by_type("http://type/a")
        assert len(found) == 2


class TestPart:
    def test_basic_properties(self) -> None:
        p = Part("/ppt/presentation.xml", "application/xml", b"<data/>")
        assert p.part_name == "/ppt/presentation.xml"
        assert p.content_type == "application/xml"
        assert p.blob == b"<data/>"
        assert len(p.rels) == 0

    def test_blob_setter(self) -> None:
        p = Part("/test.xml", "text/xml", b"old")
        p.blob = b"new"
        assert p.blob == b"new"


class TestOpcPackage:
    def test_roundtrip(self) -> None:
        pkg = OpcPackage()
        pkg.content_types.add_default("xml", "application/xml")
        pkg.content_types.add_default("rels", "application/vnd.openxmlformats-package.relationships+xml")
        pkg.rels.add("http://type/doc", "ppt/presentation.xml")
        part = Part("/ppt/presentation.xml", "application/xml", b"<root/>")
        pkg.add_part(part)

        buf = io.BytesIO()
        pkg.save(buf)
        buf.seek(0)

        pkg2 = OpcPackage.open(buf)
        assert len(pkg2.parts) == 1
        p = pkg2.get_part("/ppt/presentation.xml")
        assert p is not None
        assert p.blob == b"<root/>"

    def test_part_rels_roundtrip(self) -> None:
        pkg = OpcPackage()
        pkg.content_types.add_default("xml", "application/xml")
        part = Part("/ppt/presentation.xml", "application/xml", b"<root/>")
        part.rels.add("http://type/slide", "slides/slide1.xml")
        pkg.add_part(part)

        buf = io.BytesIO()
        pkg.save(buf)
        buf.seek(0)

        pkg2 = OpcPackage.open(buf)
        p = pkg2.get_part("/ppt/presentation.xml")
        assert p is not None
        assert len(p.rels) == 1


class TestUnits:
    def test_inches(self) -> None:
        emu = Inches(1.0)
        assert int(emu) == 914400
        assert abs(emu.inches - 1.0) < 0.001

    def test_cm(self) -> None:
        emu = Cm(2.54)
        assert abs(emu.inches - 1.0) < 0.01

    def test_pt(self) -> None:
        emu = Pt(72.0)
        assert int(emu) == 914400

    def test_emu_repr(self) -> None:
        assert repr(Emu(914400)) == "Emu(914400)"


class TestRGBColor:
    def test_create(self) -> None:
        c = RGBColor(255, 0, 128)
        assert c.r == 255
        assert c.g == 0
        assert c.b == 128
        assert str(c) == "FF0080"

    def test_from_string(self) -> None:
        c = RGBColor.from_string("FF0080")
        assert c.r == 255
        assert c.g == 0
        assert c.b == 128

    def test_from_string_with_hash(self) -> None:
        c = RGBColor.from_string("#00FF00")
        assert c == RGBColor(0, 255, 0)

    def test_invalid_range(self) -> None:
        import pytest

        with pytest.raises(ValueError):
            RGBColor(256, 0, 0)

    def test_equality(self) -> None:
        assert RGBColor(1, 2, 3) == RGBColor(1, 2, 3)
        assert RGBColor(1, 2, 3) != RGBColor(4, 5, 6)


class TestNamespaces:
    def test_qn(self) -> None:
        result = qn("p:sld")
        assert result == "{http://schemas.openxmlformats.org/presentationml/2006/main}sld"

    def test_qn_unknown(self) -> None:
        import pytest

        with pytest.raises(KeyError):
            qn("unknown:tag")
